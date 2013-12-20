<%
IF request.Form("checkOrd"&r)="YES" THEN
	pOrderStatus=request.Form("orderstatus"&r)
	pCheckEmail=request.Form("checkEmail"&r)
	pIdOrder=Request.Form("idOrder"&r)  & ""
	qry_ID=pIdOrder
	call opendb()
	
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
	Sub CreateDownloadInfo2(pIDProduct,pQuantity)
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
	IF DPOrder="1" then
		query="select idproduct,quantity,idconfigSession from ProductsOrdered WHERE idOrder="& qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
	
		do while not rs.eof
			pIdProduct=rs("idproduct")
			pQuantity=rs("quantity")
			tmpidConfig=rs("idconfigSession")
			Call CreateDownloadInfo2(pIDProduct,pQuantity)
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
									Call CreateDownloadInfo2(PrdArr(k),QtyArr(k)*pQuantity)
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
									Call CreateDownloadInfo2(CPrdArr(k),1)
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
	query="SELECT orders.idcustomer,orders.address,orders.City,orders.StateCode,orders.zip,orders.CountryCode,orders.shippingAddress,orders.shippingCity,orders.shippingStateCode,orders.shippingZip,orders.shippingCountryCode,orders.ShipmentDetails,orders.PaymentDetails,orders.discountDetails,orders.taxAmount,orders.total,orders.comments,orders.ShippingFullName,orders.address2,orders.ShippingCompany,orders.ShippingAddress2,orders.taxDetails,orders.iRewardValue,orders.iRewardRefId, orders.iRewardPointsRef,orders.iRewardPointsCustAccrued,customers.phone,ord_DeliveryDate,ord_VAT,pcOrd_CatDiscounts FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" &qry_ID
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
		pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)
		pord_VAT=rsEmailInfo("ord_VAT")
		pcOrd_CatDiscounts=rsEmailInfo("pcOrd_CatDiscounts")
		
	set rsEmailInfo=nothing
	
	'Get customer details for this order
	query="Select name,lastname,customerCompany,email FROM customers WHERE idcustomer="& pIdCustomer
	Set rsCust=Server.CreateObject("ADODB.Recordset")
	Set rsCust=conntemp.execute(query)
		pName=rsCust("name")
		pLName=rsCust("lastname")
		pCustomerCompany=rsCust("customerCompany")
		pEmail=rsCust("email")
	Set rsCust=nothing

	'Send Order Confirmation email to customer, if checked
	if pCheckEmail="YES" then%>
		<!--#include file="sendmailCustomerProcessed.asp"-->
		<% 
		pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_6")
		call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerEmail, "&quot;", chr(34)))
	end if
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
%>


<% 
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
%>