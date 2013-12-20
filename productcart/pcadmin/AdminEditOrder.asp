<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Dim pageTitle, Section
pageTitle="Edit Order"
Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<!-- #Include file="../pc/checkdate.asp" -->
<!--#include file="AdminHeader.asp"-->

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<%
'********************************************************************
'// Multiply Gift Wrapping charge times units of product purchased?
'// 1 = YES; 0 = NO
pcIntMultipleGiftWrap = 0
'********************************************************************
%>

<script language="JavaScript">
<!--

	function newWindow(file,window) {
	msgWindow=open(file,window,'resizable=yes,width=500,height=500,scrollbars=1');
	if (msgWindow.opener == null) msgWindow.opener = self;
	}

	function newWindow2(file,window) {
			catWindow=open(file,window,'resizable=no,width=480,height=360,scrollbars=1');
			if (catWindow.opener == null) catWindow.opener = self;
	}

	function win2(fileName) {
		myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=400,height=300')
		myFloater.location.href=fileName;
	}

	function validate()
	{
		var fields=document.getElementsByTagName("input");
		for (f = 0; f < fields.length; f++)
		{
			if(fields[f].type=="text" && fields[f].id !="")
			{
				var idText=fields[f].id;
				if(idText.substr(0,3)=="qty" && !isNaN(idText.substr(3)))
				{
					if (Number(document.getElementById(idText.substr(0,3)+idText.substr(3)).value)<1)
					{
						alert("Please enter a valid product quantity. You cannot set it to 0 or a negative number. If you need to, you can remove a product from your order as long as the order is still 'Pending'.");
						return false;
					}
				}
			}
		}
		return true;
	}

//-->
</script>

<%
dim conntemp, rs, query
call opendb()

qryID=request.QueryString("ido")
if qryID="" then
	qryID=request.Form("ido")
end if

query="SELECT orderstatus FROM orders WHERE idOrder="&qryID&";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
intOrderStatus=rs("orderstatus")
set rs=nothing

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: DELETE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request.QueryString("del")="YES" then
	delPrd=request.QueryString("delPrd")
	query="DELETE FROM ProductsOrdered WHERE idProductOrdered="&delPrd&" AND idOrder="&qryID&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing

	mCount=session("admin_cp_edit_count")
	For j=1 to mCount
		if clng(delPrd)=clng(session("admin_cp_" & qryID & "_idprd" & j)) then
			pcv_qty=session("admin_cp_" & qryID & "_qty" & j)
			query="UPDATE Products SET stock=stock+" & pcv_qty & ",sales=sales-" & pcv_qty & " WHERE idproduct=" & session("admin_cp_" & qryID & "_id" & j)
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
			exit for
		end if
	Next
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: DELETE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

if request.QueryString("action")="upd" OR request.Form("SubmitUPD")<>"" then
	qry_ID=qryID
	pIdOrder=qryID

	'//Add to Audit Log
	'Audit datetime
	If Session("pcAuditAdmin")&""<>"" Then
		pcAuditDateTime=CheckDateSQL(now())
		query = "INSERT INTO pcAdminAuditLog (idAdmin, idOrder, pcAdminAuditDate, pcAdminAuditPage) VALUES ("&Session("pcAuditAdmin")&", "&pIdOrder&", '"&pcAuditDateTime&"', 'AdminEditOrder');"
		set rs=connTemp.execute(query)
	End If

	query="SELECT orderStatus,idcustomer FROM Orders WHERE idorder=" & qryID & ";"
	set rs=connTemp.execute(query)
	pcv_ordStatus=0
	if not rs.eof then
		pcv_ordStatus=rs("orderStatus")
		pIdCustomer=rs("idcustomer")
	end if
	set rs=nothing

	Dim pTodaysDate
	pTodaysDate=Date()
	if SQL_Format="1" then
		pTodaysDate=Day(pTodaysDate)&"/"&Month(pTodaysDate)&"/"&Year(pTodaysDate)
	else
		pTodaysDate=Month(pTodaysDate)&"/"&Day(pTodaysDate)&"/"&Year(pTodaysDate)
	end if

	IF pcv_ordStatus="3" THEN

	pcv_OrdDPs=0
	pcv_OrdGCs=0
	query="SELECT products.idproduct,ProductsOrdered.quantity FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qryID & " AND products.Downloadable<>0;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_OrdDPs=1
	end if
	set rs=nothing

	query="SELECT products.idproduct FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qryID & " AND products.pcprod_GC<>0;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_OrdGCs=1
	end if
	set rs=nothing

	query="UPDATE Orders SET DPs=" & pcv_OrdDPs & ",pcOrd_GCs=" & pcv_OrdGCs & " WHERE idorder=" & qryID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing

	'Call License Generator for Standard & BTO Products
	if pcv_OrdDPs="1" then
		query="SELECT products.idproduct FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qryID & " AND products.Downloadable<>0 GROUP BY products.idproduct;"
		set rstemp=connTemp.execute(query)
		do while not rstemp.eof
			query="select * from Products,DProducts where products.idproduct=" & rstemp("idproduct") & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
			set rs=connTemp.execute(query)

			if not rs.eof then
				pIdproduct=rstemp("idproduct")
				pSku=rs("sku")
				query="SELECT SUM(quantity) as Tquantity FROM ProductsOrdered WHERE idorder=" & qryID & " AND idproduct=" & pIdproduct & " GROUP BY idproduct;"
				set rsQ=connTemp.execute(query)
				pQuantity=rsQ("Tquantity")
				set rsQ=nothing
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

					pcv_LTotal=0
					query="SELECT Count(*) AS Total From DPLicenses WHERE idOrder="& qry_ID & " AND IDProduct=" & pIdproduct & " GROUP BY IDProduct;"
					Set rsQ=connTemp.execute(query)
					IF not rsQ.eof then
						pcv_LTotal=rsQ("Total")
					END IF
					set rsQ=nothing

					For k=0 to Cint(pQuantity)-1-pcv_LTotal
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
						set rstemp19=connTemp.execute(query)
					Next
				end if

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
				set rstemp19=connTemp.execute(query)

				if not rstemp19.eof then
					ReqExist=1
				end if

				LOOP UNTIL ReqExist=0

				'Insert Standard & BTO Products Download Requests into DPRequests Table

				if scDB="Access" then
				query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "',#" & pTodaysDate & "#)"
				else
				query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "','" & pTodaysDate & "')"
				end if
				set rstemp19=connTemp.execute(query)
			end if
			rstemp.moveNext
		loop
	end if

	'GGG Add-on start
	IF pcv_OrdGCs=1 then
		query="SELECT products.idproduct FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qryID & " AND products.pcprod_GC<>0 GROUP BY Products.idproduct;"
		set rstemp=connTemp.execute(query)
		DO while not rstemp.eof
			query="select pcGC.pcGC_Exp,pcGC.pcGC_ExpDate,pcGC.pcGC_ExpDays,pcGC.pcGC_CodeGen,pcGC.pcGC_GenFile,products.sku,products.price from pcGC,Products where pcGC.pcGC_idproduct=" & rstemp("idproduct") & " and Products.idproduct=pcGC.pcGC_idproduct and products.pcprod_GC=1"
			set rs=connTemp.execute(query)

			if not rs.eof then
				pIdproduct=rstemp("idproduct")
				query="SELECT SUM(quantity) as Tquantity FROM ProductsOrdered WHERE idorder=" & qryID & " AND idproduct=" & pIdproduct & " GROUP BY idproduct;"
				set rsQ=connTemp.execute(query)
				pQuantity=rsQ("Tquantity")
				set rsQ=nothing
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

					pcv_GTotal=0
					query="SELECT Count(*) AS Total From pcGCOrdered WHERE pcGO_idOrder="& qry_ID & " AND pcGO_IDProduct=" & pIdproduct & ";"
					Set rsQ=connTemp.execute(query)
					IF not rsQ.eof then
						pcv_GTotal=rsQ("Total")
					END IF
					set rsQ=nothing

					For k=1 to Cint(pQuantity)-pcv_GTotal

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
					set rsG=connTemp.execute(query)

					if not rsG.eof then
					ReqExist=1
					end if

					LOOP UNTIL ReqExist=0

					'Insert Gift Codes to Database

					if scDB="Access" then
					query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "',#" & pGCExpDate & "#," & pGCAmount & ",1)"
					else
					query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "','" & pGCExpDate & "'," & pGCAmount & ",1)"
					end if
					set rsG=connTemp.execute(query)

					Next

				END IF

			end if
			rstemp.moveNext
		LOOP
	END IF
	END IF
	'GGG Add-on end

	query="SELECT products.idproduct,products.pcDropShipper_ID FROM Products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qryID & " AND products.pcDropShipper_ID<>0;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
		For i=0 to intCount
			query="UPDATE ProductsOrdered SET pcDropShipper_ID=" & pcArr(1,i) & " WHERE idproduct=" & pcArr(0,i) & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		Next
	end if
	set rs=nothing
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: UPD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request.QueryString("action")="upd" then

	query="SELECT ProductsOrdered.pcPO_GWPrice,ProductsOrdered.idProductOrdered, ProductsOrdered.idOrder, ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.unitPrice,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts,ProductsOrdered.IdConfigSession, products.description, products.weight, products.sku, products.notax FROM products INNER JOIN ProductsOrdered ON products.idProduct = ProductsOrdered.idProduct WHERE (((ProductsOrdered.idOrder)="&qryID&"));"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if rs.eof then
		'no orders match this id
		call closedb()
	end if
	pCnt=0
	subtotal=0
	taxabletotal=0
	shipHandling=0
	'GGG Add-on start
	subGW=0
	'GGG Add-on end

	do until rs.eof
		pCnt=pCnt+1
		'GGG Add-on start
		pcv_GWPrice=rs("pcPO_GWPrice")
		if pcv_GWPrice<>"" then
		else
		pcv_GWPrice=0
		end if
		subGW=subGW+pcv_GWPrice
		'GGG Add-on end
		pIdProductOrdered=rs("idProductOrdered")
		pIdOrder=rs("idOrder")
		pIdProduct=rs("idProduct")
		pQuantity=rs("quantity")
		tQTY=pQuantity
		pOnlinePrice=rs("unitPrice")
		pQDiscounts=rs("QDiscounts")
		if isNull(pQDiscounts) or pQDiscounts=""  then
			pQDiscounts=0
		end if
		pItemsDiscounts=rs("ItemsDiscounts")
		if isNull(pItemsDiscounts) or pItemsDiscounts=""  then
			pItemsDiscounts=0
		end if
		pIdConfigSession=rs("IdConfigSession")
		pNoTax=rs("notax")

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: ReCalculate BTO Items Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		itemsDiscounts=0
		if scBTO=1 then
		if pIdConfigSession<>"0" then
			query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsQ=conntemp.execute(query)
			stringProducts=rsQ("stringProducts")
			stringValues=rsQ("stringValues")
			stringCategories=rsQ("stringCategories")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			Qstring=rsQ("stringQuantity")
			ArrQuantity=Split(Qstring,",")
			Pstring=rsQ("stringPrice")
			ArrPrice=split(Pstring,",")

			if ArrProduct(0)<>"na" then
				for j=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(j)
					set rs99=connTemp.execute(query)
					TempDiscount=0
					do while not rs99.eof
						QFrom=rs99("quantityFrom")
						QTo=rs99("quantityUntil")
						DUnit=rs99("discountperUnit")
						QPercent=rs99("percentage")
						DWUnit=rs99("discountperWUnit")
						if (DWUnit=0) and (DUnit>0) then
							DWUnit=DUnit
						end if


						TempD1=0
						if (clng(ArrQuantity(j)*tQTY)>=clng(QFrom)) and (clng(ArrQuantity(j)*tQTY)<=clng(QTo)) then
							if QPercent="-1" then
								if session("customerType")=1 then
									TempD1=ArrQuantity(j)*tQTY*ArrPrice(j)*0.01*DWUnit
								else
									TempD1=ArrQuantity(j)*tQTY*ArrPrice(j)*0.01*DUnit
								end if
							else
								if session("customerType")=1 then
									TempD1=ArrQuantity(j)*tQTY*DWUnit
								else
									TempD1=ArrQuantity(j)*tQTY*DUnit
								end if
							end if
						end if
						TempDiscount=TempDiscount+TempD1
						rs99.movenext
					loop
					itemsDiscounts=ItemsDiscounts+TempDiscount
				next
			end if 'Have BTO Items
		end if 'Have ConfigSession
		end if
		pItemsDiscounts=itemsDiscounts

		query="UPDATE ProductsOrdered SET ItemsDiscounts=" & replacecommaToDB(pItemsDiscounts) & " WHERE idProductOrdered="&pIdProductOrdered&" AND idOrder="&pIdOrder&";"
		set rsQ=server.CreateObject("ADODB.RecordSet")
		set rsQ=connTemp.execute(query)
		set rsQ=nothing

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: ReCalculate BTO Items Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		'BTO Additional Charges
		Charges=0
		if scBTO=1 then
			if pIdConfigSession<>"0" then
				query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
				set rsConfigObj=conntemp.execute(query)
				if err.number <> 0 then
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails: "&err.description)
				end if
				stringCProducts=rsConfigObj("stringCProducts")
				stringCValues=rsConfigObj("stringCValues")
				stringCCategories=rsConfigObj("stringCCategories")
				ArrCProduct=Split(stringCProducts, ",")
				ArrCValue=Split(stringCValues, ",")
				ArrCCategory=Split(stringCCategories, ",")
				if ArrCProduct(0)<>"na" then
					for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
						query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))"
						set rsConfigObj=connTemp.execute(query)
						if (CDbl(ArrCValue(i))>0)then
							Charges=Charges+cdbl(ArrCValue(i))
						end if
						set rsConfigObj=nothing
					next
					set rsConfigObj=nothing
				end if
			end if
		end if 'BTO Additional Charges

		subtotal=subtotal+(cdbl(pOnlinePrice)*Clng(pQuantity))-pItemsDiscounts-pQDiscounts+Charges

		if pNoTax=0 then
			taxabletotal=taxabletotal+(cdbl(pOnlinePrice)*int(pQuantity))-pItemsDiscounts-pQDiscounts+Charges
		end if
		rs.moveNext
	loop

	query="SELECT idCustomer, total, taxAmount, shipmentDetails, paymentDetails, discountDetails, iRewardPoints, iRewardValue,pcOrd_CatDiscounts FROM orders WHERE idOrder="&qryID&";"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pIdCustomer=rs("idCustomer")
	ptotal=rs("total")
	ptaxAmount=rs("taxAmount")
	shipmentArray=rs("shipmentDetails")
	paymentArray=rs("paymentDetails")
	discountArray=rs("discountDetails")
	piRewardPoints=rs("iRewardPoints")
	pIRewardValue=rs("iRewardValue")
	CatDiscounts=rs("pcOrd_CatDiscounts")
	if CatDiscounts<>"" then
	else
	CatDiscounts="0"
	end if

	shipSplit=split(shipmentArray,",")
	varShip="1"
	if ubound(shipSplit)>1 then
		if NOT isNumeric(trim(shipSplit(2))) then
			varShip="0"
		else
			shipProvider=shipSplit(0)
			shipService=shipSplit(1)
			shipPrice=trim(shipSplit(2))
			shipPrice=replace(shipPrice,CHR(13),"")
			if ubound(shipSplit)=>3 then
				shipHandling=trim(shipSplit(3))
				if NOT isNumeric(shipHandling) then
					shipHandling=0
				end if
			else
				shipHandling=0
			end if
			if ubound(shipSplit)>4 then
				shipServiceCode=trim(shipSplit(5))
			end if
		end if
	else
		varShip="0"
		shipPrice=0
	end if
	paymentSplit=split(paymentArray,"||")
	paymentType=paymentSplit(0)
	if ubound(paymentSplit)>0 then
		paymentPrice=trim(paymentSplit(1))
		if NOT isNumeric(paymentPrice) then
			paymentPrice=0
		end if
	else
		paymentPrice=0
	end if

	'=====================
	if instr(discountArray,",") then
		DiscountDetailsArry=split(discountArray,",")
		intArryCnt=ubound(DiscountDetailsArry)
	else
		intArryCnt=0
	end if
	discountTotalPrice=0
	discountTotalPriceTax=0
	for k=0 to intArryCnt
		if intArryCnt=0 then
			pTempDiscountDetails=discountArray
		else
			pTempDiscountDetails=DiscountDetailsArry(k)
		end if
		if instr(pTempDiscountDetails,"- ||") then
			discounts = split(pTempDiscountDetails,"- ||")
			discountType = discounts(0)
			discountPrice = discounts(1)
			tmpFreeShip=0
			if InStr(discountType,"(")>0 AND InStr(discountType,")")>0 then
				tmp1=split(discountType,"(")
				tmp2=split(tmp1(1),")")
				tmpDiscountCode=tmp2(0)
				query="SELECT pcDFShip.pcFShip_IDShipOpt FROM pcDFShip INNER JOIN discounts ON pcDFShip.pcFShip_IDDiscount=discounts.iddiscount WHERE discounts.discountcode like '" & tmpDiscountCode & "';"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if not rs.eof then
					tmpFreeShip=1
				end if
			end if
			if (tmpFreeShip=0) OR ((tmpFreeShip=1) AND (TaxOnCharges=1)) then
				discountTotalPriceTax=discountTotalPriceTax+discountPrice
			end if
			discountTotalPrice=discountTotalPrice+discountPrice
		end if
	Next

	tDiscountString=discountArray

	tSubTotal=subtotal

	subtotal=subtotal - discountTotalPrice - cdbl(CatDiscounts)
	if subtotal<0 then
		subtotal=0
		discountTotalPrice=tSubTotal
	end if

	'discounts affect the tax total
	taxabletotal=taxabletotal-discountTotalPriceTax

	subtotal=subtotal+paymentPrice
	taxabletotal=taxabletotal+paymentPrice
	'--------------------
	'// Start Reward Points
	useRewards=trim(request.queryString("iRewardPoints"))
	if useRewards="" then
		useRewards=0
	end if
	useRewards=int(useRewards)+int(piRewardPoints)
	If RewardsActive=1 And int(useRewards) > 0 Then
		iDollarValue=useRewards * (RewardsPercent / 100)
		if subtotal<>0 then
			subtotal=subtotal - iDollarValue
		else
			subtotal=0
		end if
		taxabletotal=taxabletotal-iDollarValue
		if taxabletotal<0 then
			taxabletotal=0
		end if
		if subtotal<0 then
			xVar=(subtotal+iDollarValue)/(RewardsPercent/100)
			useRewards=Round(xVar)
			iDollarValue=useRewards * (RewardsPercent / 100)
			subtotal=0
		end if
	Else
		iDollarValue=0
	End If

	'// End Reward Points
	subtotal=Round(subtotal,2)

	subtotal=subtotal+shipPrice
	if TaxOnCharges=1 then
		taxabletotal=taxabletotal+shipPrice
	end if

	'are there handling charges?
	subtotal=subtotal+shipHandling
	if TaxOnFees=1 then
		taxabletotal=taxabletotal+shipHandling
	end if

	'--------------------
	'calculate everything and UPdate database
	tTotal=subtotal+ptaxAmount
	if varShip="0" then
		tShippingString=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	else
		tShippingString=shipProvider&","&shipService&","&shipPrice&","&shipHandling
	end if
	tPaymentString=paymentType&"||"&paymentPrice
	'GGG Add-on start
	tTotal=tTotal+subGW
	'GGG Add-on end

	'GGG Add-on start
	'Update Gift Certificate Amount
	query="select pcOrd_GCDetails,pcOrd_GCAmount,pcOrd_GcCode,pcOrd_GcUsed from Orders WHERE idOrder="&qryID&";"
	set rs=connTemp.execute(query)
	pGiftCode=rs("pcOrd_GcCode")
	pGiftUsed=rs("pcOrd_GcUsed")
	if pGiftUsed<>"" then
	else
		pGiftUsed=rs("pcOrd_GCAmount")
		if pGiftUsed<>"" then
		else
			pGiftUsed=0
		end if
	end if
	
	tTotal=tTotal-pGiftUsed
	'GGG Add-on end
	'Update Order Details File
	pDetails=""
	query="SELECT products.description,products.sku,((ProductsOrdered.unitPrice*ProductsOrdered.quantity)) AS IAmount,ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_OptionsArray FROM products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idOrder="&qryID&";"
	set rs=ConnTemp.execute(query)
	Do while not rs.eof
		pcv_PrdName=rs("description")
		pcv_PrdSKU=rs("sku")
		pcv_PrdAmount=rs("IAmount")
		pcv_PrdQty=rs("quantity")
		pcv_strOptionsArray=rs("pcPrdOrd_OptionsArray")
		pDetails = pDetails & "  Amount: ||"& pcv_PrdAmount & " Qty:" &pcv_PrdQty& "  SKU #:" &pcv_PrdSKU & " - " &pcv_PrdName& " " & " " & pcv_strOptionsArray & Vbcrlf

		rs.MoveNext
	Loop
	set rs=nothing
	if pDetails<>"" then
		pDetails = replace(pDetails,"'","''")
		pDetails=replace(pDetails,"''''","''")
	end if
	query="UPDATE orders SET details='" & pDetails & "',total="&tTotal&",taxAmount="&ptaxAmount&",shipmentDetails='"&replace(tShippingString,"'","''")&"',paymentDetails='"&replace(tPaymentString,"'","''")&"',discountDetails='"&replace(tDiscountString,"'","''")&"',iRewardPoints="&useRewards&", iRewardValue="&iDollarValue&",pcOrd_GcUsed=" & pGiftUsed & ",pcOrd_GWTotal=" & subGW & " WHERE idOrder="&qryID&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing

	query="SELECT idProductOrdered,idProduct,quantity FROM ProductsOrdered WHERE idOrder="&qryID&";"
	set rs=connTemp.execute(query)

	do while not rs.eof
		pidProductOrdered=rs("idProductOrdered")
		pidProduct=rs("idProduct")
		newqty=rs("quantity")
		mCount=session("admin_cp_edit_count")
		haveit=0
		For j=1 to mCount
			if clng(pidProductOrdered)=clng(session("admin_cp_" & qryID & "_idprd" & j)) then
				haveit=1
				pcv_qty=session("admin_cp_" & qryID & "_qty" & j)
				newqty=clng(newqty)-clng(pcv_qty)
				query="UPDATE Products SET stock=stock-" & newqty & ",sales=sales+" & newqty & " WHERE idproduct=" & pidProduct
				set rs1=server.CreateObject("ADODB.RecordSet")
				set rs1=conntemp.execute(query)
				set rs1=nothing
				exit for
			end if
		Next
		if haveit=0 then
			query="UPDATE Products SET stock=stock-" & newqty & ",sales=sales+" & newqty & " WHERE idproduct=" & pidProduct
			set rs1=server.CreateObject("ADODB.RecordSet")
			set rs1=conntemp.execute(query)
			set rs1=nothing
		end if
		rs.MoveNext

	loop
	set rs=nothing

	'*** UPDATE Order and Customer Reward Points ***

	'// 1. Load Reward Points (RP) currently associated with this order
	query="SELECT iRewardPointsCustAccrued FROM orders WHERE idorder=" & qryID & ";"
	set rsQ=connTemp.execute(query)
	preRewards=0
	if not rsQ.eof then
		preRewards=cdbl(rsQ("iRewardPointsCustAccrued"))
	end if
	set rsQ=nothing

	'// 2. Recalculate RP for this order
	query="SELECT products.iRewardPoints*ProductsOrdered.quantity AS pdiscount FROM products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & qryID & ";"
	set rsQ=connTemp.execute(query)
	curRewardPointsAccrued=0
	do while not rsQ.eof
		curRewardPointsAccrued=curRewardPointsAccrued+cdbl(rsQ("pdiscount"))
		rsQ.MoveNext
	loop
	set rsQ=nothing

	'// 3. Determine the difference between the current RP and the original RP
	diffRewardPointsAccrued=cdbl(curRewardPointsAccrued)-Cdbl(preRewards)

	'// 4. If the order has already been processed, apply the difference to the pointes accrued by the customer
	if intOrderStatus > 2 then
		query="UPDATE customers SET iRewardPointsAccrued=iRewardPointsAccrued+" & diffRewardPointsAccrued & " WHERE idCustomer="&pIdCustomer&";"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if

	'// 5. Update the order with the new RP value calculated above
	query="UPDATE orders SET iRewardPointsCustAccrued=" & curRewardPointsAccrued & " WHERE idorder=" & qryID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing


	'*** END OF UPDATE REWARD POINTS ***

	call closedb()
	if request.QueryString("del")="YES" then
		response.redirect "AdminEditOrder.asp?action=upds&ido="&qryID&"&removed=yes"
	else
		response.redirect "AdminEditOrder.asp?action=upds&ido="&qryID
	end if
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: UPD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: POST BACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request.Form("SubmitUPD")<>"" then
	tCnt=request.form("tCnt")
	subtotal=0
	taxabletotal=0
	CatDiscTotal=0
	'GGG Add-on start
	subGW=0
	'GGG Add-on end


	'*****************************************************
	'// START: LOOP AND RECALCULATE EACH PRODUCT ORDERED
	'*****************************************************
	for i=1 to tCnt
		pcv_strActualPrice=0
		tIdProductOrdered=request.Form("IdProductOrdered"&i)
		tIdProduct=request.Form("IdProduct"&i)
		tIdOrder=request.Form("IdOrder"&i)

		'// Request and validate Product Quantity
		tQTY=request.Form("QTY"&i)
			if (not validNum(tQTY)) then
				call closeDb()
				response.Redirect "AdminEditOrder.asp?ido="&qryID&"&msg=" & server.URLEncode("You did not enter a valid quantity for one or more products.")
			else
				tQTY=tQTY
			end if

		pcv_strOptionsPriceArray=request.Form("strOptionsPriceArray"&i)
		toptionsLineTotal=request.Form("optionsLineTotal"&i)

		'// Request and validate Unit Price
		tunitPrice=request.Form("unitPrice"&i)
			if not isNumeric(tunitPrice) then
				call closeDb()
				response.Redirect "AdminEditOrder.asp?ido="&qryID&"&msg=" & server.URLEncode("You did not enter a valid unit price for one or more products.")
			else
				tunitPrice=replacecommaToCal(tunitPrice)
			end if

		'// Request Saved Unit Price. No validation necessary as there is no input for it.
		tSavedUnitPrice=replacecommaToCal(request.Form("SavedUnitPrice"&i))

		'// If the Price was edited set a flag. Use the actual price without re-calculating.
		if tunitPrice <> tSavedUnitPrice then
			pcv_strActualPrice=1
		end if
		tunitPrice=tunitPrice+toptionsLineTotal
		ttaxProduct=request.Form("taxProduct"&i)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: GGG Add-on and Options Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		Dim pcv_GWPrice, subGW
		pcv_GWPrice = 0
		subGW = 0
		query="SELECT pcPO_GWPrice, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray FROM ProductsOrdered WHERE idProductOrdered=" & tIdProductOrdered
		set rsGW=connTemp.execute(query)
		pcv_GWPrice=rsGW("pcPO_GWPrice")
		pcv_strSelectedOptions=rsGW("pcPrdOrd_SelectedOptions")
		pcv_strOptionsPriceArray=rsGW("pcPrdOrd_OptionsPriceArray")
		if pcv_GWPrice<>"" then
		else
		pcv_GWPrice=0
		end if
		subGW=subGW+pcv_GWPrice

		'// Gift Wrapping Total
		if pcIntMultipleGiftWrap <> 0 then
			subGWTotal = subGWTotal + (subGW*cdbl(tQTY))
		else
			subGWTotal = subGWTotal + subGW
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: GGG Add-on and Options Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~





		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Get the original Quantity Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="SELECT QDiscounts FROM ProductsOrdered WHERE idProductOrdered=" & tIdProductOrdered
		set rsPricing=server.CreateObject("ADODB.RecordSet")
		set rsPricing=connTemp.execute(query)
		tempQDiscounts=0
		'// If there is a QDiscount and the price is not actual.
		if not rsPricing.eof AND pcv_strActualPrice=0 then
			tempQDiscounts=rsPricing("QDiscounts")
		else
			tempQDiscounts=0
		end if
		set rsPricing = nothing
		if tempQDiscounts="" OR isNULL(tempQDiscounts)=True then
			tempQDiscounts=0
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Get the original Quantity Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: ReCalculate BTO Items Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		itemsDiscounts=0
		query="select idconfigSession,unitPrice from ProductsOrdered where IdProductOrdered=" & tIdProductOrdered
		set rs=connTemp.execute(query)
		pIdConfigSession="0"
		if not rs.eof then
			pIdConfigSession=rs("idconfigSession")
		end if
		if pIdConfigSession<>"0" then
			query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rs=conntemp.execute(query)
			stringProducts=rs("stringProducts")
			stringValues=rs("stringValues")
			stringCategories=rs("stringCategories")
			ArrProduct=Split(stringProducts, ",")
			ArrValue=Split(stringValues, ",")
			ArrCategory=Split(stringCategories, ",")
			Qstring=rs("stringQuantity")
			ArrQuantity=Split(Qstring,",")
			Pstring=rs("stringPrice")
			ArrPrice=split(Pstring,",")

			if ArrProduct(0)<>"na" then
				for j=lbound(ArrProduct) to (UBound(ArrProduct)-1)
					query="select * from discountsPerQuantity where IDProduct=" & ArrProduct(j)
					set rs99=connTemp.execute(query)
					TempDiscount=0
					do while not rs99.eof
						QFrom=rs99("quantityFrom")
						QTo=rs99("quantityUntil")
						DUnit=rs99("discountperUnit")
						QPercent=rs99("percentage")
						DWUnit=rs99("discountperWUnit")
						if (DWUnit=0) and (DUnit>0) then
							DWUnit=DUnit
						end if


						TempD1=0
						if (clng(ArrQuantity(j)*tQTY)>=clng(QFrom)) and (clng(ArrQuantity(j)*tQTY)<=clng(QTo)) then
							if QPercent="-1" then
								if session("customerType")=1 then
									TempD1=ArrQuantity(j)*tQTY*ArrPrice(j)*0.01*DWUnit
								else
									TempD1=ArrQuantity(j)*tQTY*ArrPrice(j)*0.01*DUnit
								end if
							else
								if session("customerType")=1 then
									TempD1=ArrQuantity(j)*tQTY*DWUnit
								else
									TempD1=ArrQuantity(j)*tQTY*DUnit
								end if
							end if
						end if
						TempDiscount=TempDiscount+TempD1
						rs99.movenext
					loop
					itemsDiscounts=ItemsDiscounts+TempDiscount
				next
			end if 'Have BTO Items
		end if 'Have ConfigSession
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: ReCalculate BTO Items Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: BTO Additional Charges
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		Charges=0
		if pIdConfigSession<>"0" then
			query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
			set rsConfigObj=conntemp.execute(query)
			if err.number <> 0 then
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails: "&err.description)
			end if
			stringCProducts=rsConfigObj("stringCProducts")
			stringCValues=rsConfigObj("stringCValues")
			stringCCategories=rsConfigObj("stringCCategories")
			ArrCProduct=Split(stringCProducts, ",")
			ArrCValue=Split(stringCValues, ",")
			ArrCCategory=Split(stringCCategories, ",")
			if ArrCProduct(0)<>"na" then
				for j=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
					query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(j)&") AND ((products.idProduct)="&ArrCProduct(j)&"))"
					set rsConfigObj=connTemp.execute(query)
					if (CDbl(ArrCValue(j))>0)then
						Charges=Charges+cdbl(ArrCValue(j))
					end if
					set rsConfigObj=nothing
				next
				set rsConfigObj=nothing
			end if
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: BTO Additional Charges
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: ReCalculate Product Quantity Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pOrigPrice=(request.Form("OriginalPrice"&i))
		if pOrigPrice<>"" and isNumeric(pOrigPrice) then
			'// add on a fraction to help vbscript round function up to the actual price
			pOrigPrice=round(request.Form("OriginalPrice"&i), 2)
		end if

		query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &tIdProduct& " AND quantityFrom<=" &tQTY& " AND quantityUntil>=" &tQTY
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)
		tempNum=0

		pOrigPrice = (pOrigPrice*cdbl(tQTY)) - itemsDiscounts


		if pIdConfigSession<>"0" then
			pOrigPrice = (tunitPrice*cdbl(tQTY)) - itemsDiscounts
		end if

		'// Exclude options from the price
		pOrigPriceNoOptions = pOrigPrice
		if isNULL(toptionsLineTotal)=False and len(toptionsLineTotal)>0 then	'// there is pricing on options
			pOrigPriceNoOptions=(pOrigPrice-(toptionsLineTotal*cdbl(tQTY)))
		end if

		pta_Price=pOrigPrice
		if not rstemp.eof and err.number<>9 then
			'// There are quantity discounts defined for that quantity
			pDiscountPerUnit=rstemp("discountPerUnit")
			pDiscountPerWUnit=rstemp("discountPerWUnit")
			pPercentage=rstemp("percentage")
			pbaseproductonly=rstemp("baseproductonly")
			if pbaseproductonly="-1" then
				pOrigPrice=pOrigPriceNoOptions
			end if
			if session("customerType")<>1 then
				if pPercentage="0" then
					if pIdConfigSession<>"0" then
						pta_Price=pta_Price - pDiscountPerUnit
					else
						pta_Price=pta_Price - (pDiscountPerUnit * tQTY)
					end if
					tempNum=tempNum + (pDiscountPerUnit * tQTY)
				else
					pta_Price=pta_Price - ((pDiscountPerUnit/100) * pOrigPrice)
					tempNum=tempNum + ((pDiscountPerUnit/100) * pOrigPrice)
				end if
			else
				if pPercentage="0" then
					if pIdConfigSession<>"0" then
						pta_Price=pta_Price - pDiscountPerWUnit
					else
						pta_Price=pta_Price - (pDiscountPerWUnit * tQTY)
					end if
					tempNum=tempNum + (pDiscountPerWUnit * tQTY)
				else
					pta_Price=pta_Price - ((pDiscountPerWUnit/100) * pOrigPrice)
					tempNum=tempNum + ((pDiscountPerWUnit/100) * pOrigPrice)
				end if
			end if
		end if
		if TempNum="" OR isNULL(TempNum)=True then
			TempNum=0
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: ReCalculate Product Quantity Discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Calculate Final Unit Price
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if pcv_strActualPrice=0 then
			if pIdConfigSession<>"0" then
				tunitPrice = tunitPrice
			else
				'tunitPrice = (pta_Price/tQTY)
			end if
		else
			tunitPrice = (tunitPrice+tempQDiscounts)
		end if
		if tunitPrice="" OR isNULL(tunitPrice)=True then
			tunitPrice=0
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Calculate Final Unit Price
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Save unit pricing to ProductsOrdered table
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="UPDATE ProductsOrdered SET quantity="&tQTY&", unitPrice="&replacecommaToDB(tunitPrice)&",QDiscounts=" & replacecommaToDB(TempNum) & ",ItemsDiscounts=" & replacecommaToDB(ItemsDiscounts) & " WHERE idProductOrdered="&tIdProductOrdered&" AND idOrder="&tIdOrder&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Save unit pricing to ProductsOrdered table
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



		'// Cat Discounts Moved Below



		subtotal=subtotal+(tunitPrice*int(tQTY))-TempNum-itemsDiscounts+Charges

		pcv_intTotalQtyDiscounts = pcv_intTotalQtyDiscounts + TempNum

		if ttaxProduct="YES" then
			taxabletotal=taxabletotal+(tunitPrice*int(tQTY))
		end if

	next
	'*****************************************************
	'// END: LOOP AND RECALCULATE EACH PRODUCT ORDERED
	'*****************************************************


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Category-based quantity discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Discounts by Categories
		Dim pcv_strApplicableProducts
		CatDiscTotal=0
		query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		ProdUsedFlag = False
		Do While not rs.eof
			CatSubQty=0
			CatSubTotal=0
			CatSubDiscount=0
			ApplicableCategoryID = rs("IDCat")

			for o=1 to request.form("tCnt")

					query="select idproduct from categories_products where idcategory=" & rs("IDCat") & " and idproduct=" & tIdProduct
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=connTemp.execute(query)
					if not rstemp.eof then
						CatSubQty=CatSubQty + request.Form("QTY"&o)
						CatSubTotal=CatSubTotal+(request.Form("unitPrice"&o) * request.Form("QTY"&o))-itemsDiscounts
						pcv_strApplicableProducts = pcv_strApplicableProducts & tIdProduct & chr(124) &  ApplicableCategoryID & ","
					end if
					set rstemp=nothing
			Next

			pcv_strrApplicableCategories = pcv_strrApplicableCategories & CatSubTotal & chr(124) &  ApplicableCategoryID & ","

			if CatSubQty>0 then
				query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rs("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				if not rstemp.eof then
					'// There are quantity discounts defined for that quantity
					pDiscountPerUnit=rstemp("pcCD_discountPerUnit")
					pDiscountPerWUnit=rstemp("pcCD_discountPerWUnit")
					pPercentage=rstemp("pcCD_percentage")
					pbaseproductonly=rstemp("pcCD_baseproductonly")

					if session("customerType")<>1 then  'customer is a normal user
						if pPercentage="0" then
							CatSubDiscount=pDiscountPerUnit*CatSubQty
						else
							CatSubDiscount=(pDiscountPerUnit/100) * CatSubTotal
						end if
					else  'customer is a wholesale customer
						if pPercentage="0" then
							CatSubDiscount=pDiscountPerWUnit*CatSubQty
						else
							CatSubDiscount=(pDiscountPerWUnit/100) * CatSubTotal
						end if
					end if
					ProdUsedFlag = True
				end if
				set rstemp=nothing
			end if '// if CatSubQty>0 then

			CatDiscTotal=CatDiscTotal+CatSubDiscount
			rs.MoveNext
		loop
		set rs=nothing
		'// Round the Category Discount to two decimals
		if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
			CatDiscTotal = RoundTo(CatDiscTotal,.01)
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Category-based quantity discounts
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


	'*****************************************************
	'// END: LOOP AND RECALCULATE EACH PRODUCT ORDERED
	'*****************************************************





	'*****************************************************
	'// START: HANDLE DISCOUNTS
	'*****************************************************

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Generate Discount Array
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		tDiscountString=""
		tdiscounts=0
		tdiscountsTax=0
		'// Get total existing
		intArryCnt=request.Form("intArryCnt")
		'// Loop through existing
		if intArryCnt>-1 then
			for i=0 to intArryCnt
				tempDiscountType=request.Form("DiscountType"&i)
				tempDiscountPrice=request.Form("DiscountPrice"&i)
				tempDiscountPrice=replacecomma(tempDiscountPrice)
				if tempDiscountPrice<>"" AND tempDiscountPrice<>0  AND tempDiscountType<>"" then
					if tDiscountString="" then
						tDiscountString=tDiscountString&tempDiscountType&"- ||"&tempDiscountPrice
					else
						tDiscountString=tDiscountString&","&tempDiscountType&"- ||"&tempDiscountPrice
					end if
					tmpFreeShip=0
					if InStr(tempDiscountType,"(")>0 AND InStr(tempDiscountType,")")>0 then
						tmp1=split(tempDiscountType,"(")
						tmp2=split(tmp1(1),")")
						tmpDiscountCode=tmp2(0)
						query="SELECT pcDFShip.pcFShip_IDShipOpt FROM pcDFShip INNER JOIN discounts ON pcDFShip.pcFShip_IDDiscount=discounts.iddiscount WHERE discounts.discountcode like '" & tmpDiscountCode & "';"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
						if not rs.eof then
							tmpFreeShip=1
						end if
					end if
					ttaxShipping=request.Form("taxShipping")
					if (tmpFreeShip=0) OR ((tmpFreeShip=1) AND (ttaxShipping="YES")) then
						tdiscountsTax=tdiscountsTax+tempDiscountPrice
					end if
					tdiscounts=tdiscounts+tempDiscountPrice
				end if

			next
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Generate Discount Array
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Discount Calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		'// Get New Percent Discount
		tcalculateDisPercentage=trim(request.Form("calculateDisPercentageM"))

		'// Get New Discount Name
		tDiscountTypeM=request.Form("discountTypeM")
		if tDiscountTypeM = "" then
			tDiscountTypeM = "Special Discount"
		end if

		if tcalculateDisPercentage<>"" AND isNumeric(tcalculateDisPercentage) AND tcalculateDisPercentage<>"0" then
			session("adminDisPercentage")=tcalculateDisPercentage
			tdiscountM=subtotal*(tcalculateDisPercentage/100)
		else
			'// Get the flat rate discount
			tdiscountM=request.Form("discountsM")
			tdiscountM=replacecomma(tdiscountM)
		end if

		'// Append to our string of Discounts
		if tdiscountM<>"" AND tdiscountM<>"0" AND tDiscountTypeM<>"" then
			if tDiscountString="" then
				tDiscountString=tDiscountString&tDiscountTypeM&"- ||"&tdiscountM
			else
				tDiscountString=tDiscountString&","&tDiscountTypeM&"- ||"&tdiscountM
			end if
		end if

		if tdiscountM="" then
			tdiscountM=0
		end if
		tdiscounts=tdiscounts+tdiscountM
		tdiscountsTax=tdiscountsTax+tdiscountM

		if tDiscountString="" then
			tDiscountString=dictLanguage.Item(Session("language")&"_saveorder_10")
		end if

		tSubTotal=subtotal


		CatDiscounts=CatDiscTotal
		if CatDiscounts<>"" then
		else
		CatDiscounts=0
		end if

		subtotal=subtotal - tdiscounts - CatDiscounts
		if subtotal<0 then
			subtotal=0
			tdiscounts=tSubTotal
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END: Discount Calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


	'*****************************************************
	'// END: HANDLE DISCOUNTS
	'*****************************************************



	'// Discounts affect the tax total
	taxabletotal = taxabletotal - pcv_intTotalQtyDiscounts - tdiscountsTax - CatDiscounts - itemsDiscounts + Charges

	'--------------------
	'are there payment charges?
	tcalculatePayPercentage=request.Form("calculatePayPercentage")
	if tcalculatePayPercentage<>"" AND isNumeric(tcalculatePayPercentage) AND tcalculatePayPercentage<>"0" then
		session("adminPayPercentage")=tcalculatePayPercentage
		tpaymentCharges=subtotal*(tcalculatePayPercentage/100)
	else
		tpaymentCharges=request.Form("paymentCharges")
	end if
	if NOT tpaymentCharges="" then
		tpaymentCharges=replacecommaToCal(tpaymentCharges)
		subtotal=subtotal+tpaymentCharges
	else
		tpaymentCharges=0
	end if
	subtotal=Round(subtotal,2)
	taxabletotal=taxabletotal+tpaymentCharges
	'--------------------

	'// Start Reward Points
	useRewards=request.Form("iRewardPoints")
	if not validNum(useRewards) then
		useRewards=0
	end if
	If RewardsActive=1 And int(useRewards) > 0 Then
		iDollarValue=useRewards * (RewardsPercent / 100)
		if subtotal<>0 then
			subtotal=subtotal - iDollarValue
		else
			subtotal=0

		end if
		taxabletotal=taxabletotal-iDollarValue
		if taxabletotal<0 then
			taxabletotal=0
		end if
		if subtotal<0 then
			xVar=(subtotal+iDollarValue)/(RewardsPercent/100)
			useRewards=Round(xVar)
			iDollarValue=useRewards * (RewardsPercent / 100)
			subtotal=0
		end if
	Else
		iDollarValue=0
	End If

	dim X

	X=useRewards-int(request.Form("curUsedPointsOrder"))
	if not validNum(X) then
		X=0
	end if
	customerUsedTotal=int(request.Form("totalCustUsed"))+int(X)
	Y=trim(request.form("iRewardPointsCustAccrued"))
	if not IsNumeric(Y) then
		Y=0
	else
		Y=Round(Y)
	end if
	query="SELECT iRewardPointsCustAccrued FROM orders WHERE idorder=" & tIDOrder & ";"
	set rsQ=connTemp.execute(query)
	preRewards=0
	if not rsQ.eof then
		preRewards=cdbl(rsQ("iRewardPointsCustAccrued"))
	end if
	set rsQ=nothing
	query="SELECT products.iRewardPoints*ProductsOrdered.quantity AS pdiscount FROM products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & tIDOrder & ";"
	set rsQ=connTemp.execute(query)
	curRewardPointsAccrued=0
	do while not rsQ.eof
		curRewardPointsAccrued=curRewardPointsAccrued+cdbl(rsQ("pdiscount"))
		rsQ.MoveNext
	loop
	set rsQ=nothing

	curRewardPointsAccrued=trim(request.Form("curRewardPointsAccruedOrder"))
	if not IsNumeric(curRewardPointsAccrued) then
		curRewardPointsAccrued=0
	else
		curRewardPointsAccrued=Round(curRewardPointsAccrued)
	end if

	curRewardPointsAccrued=curRewardPointsAccrued+int(Y)

	'int(request.Form("curRewardPointsAccruedOrder"))
	customerAccruedTotal=cdbl(curRewardPointsAccrued)-Cdbl(preRewards)

	'int(request.Form("totalCustAccrued"))
	'// End Reward Points

	'adjust customer's reward points if necessary
	if (x<>0) or (y<>0) or (customerAccruedTotal<>0) then
		pIdCustomer=request.Form("idCustomer")
		query="SELECT customers.iRewardPointsAccrued, customers.iRewardPointsUsed FROM customers WHERE idCustomer="&pIdCustomer&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		iRewardPointsUsed=rs("iRewardPointsUsed")
		iRewardPointsUsed=iRewardPointsUsed+X
			if intOrderStatus > 2 then 'if the order is still pending, don't update the accrued Reward Points
				query="UPDATE customers SET iRewardPointsAccrued=iRewardPointsAccrued+" & customerAccruedTotal &" WHERE idCustomer="&pIdCustomer&";"
				set rs=connTemp.execute(query)
			end if
		query="UPDATE customers SET iRewardPointsUsed="&iRewardPointsUsed&" WHERE idCustomer="&pIdCustomer&";"
		set rs=connTemp.execute(query)
	end if


	'--------------------
	'are there shipping rates?
	tShipping=request.Form("Shipping")
	if tShipping="" then
		tShipping=0
	end if
	tShipping=replacecommaToCal(tShipping)
	subtotal=subtotal+tShipping
	ttaxShipping=request.Form("taxShipping")
	if ttaxShipping="YES" then
		session("adminTaxonCharges")=1
		taxabletotal=taxabletotal+tShipping
	else
		session("adminTaxonCharges")=0
	end if
	'--------------------

	'--------------------
	'are there handling charges?
	tHandling=request.Form("handling")
	if tHandling="" then
		tHandling=0
	end if
	tHandling=replacecommaToCal(tHandling)
	subtotal=subtotal+tHandling
	ttaxHandling=request.Form("taxHandling")
	if ttaxHandling="YES" then
		session("adminTaxonFees")=1
		taxabletotal=taxabletotal+tHandling
	else
		session("adminTaxonFees")=0
	end if
	'--------------------



	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Calculate Taxes
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	tTax=0
	if request.Form("splitTaxType")=1 then
		ptaxTypeCnt=request.Form("taxTypeCnt")
		tTaxAdding=0
		ptaxDetailString=""
		for i=1 to ptaxTypeCnt
			tTaxPercentage=request.Form("calculateTaxPercentage"&i)
			if tTaxPercentage="" then
			tTaxPercentage="0"
			end if
			if tTaxPercentage<>"" AND isNumeric(tTaxPercentage) AND tTaxPercentage<>"0" then
				session("adminTaxPercentage"&i)=tTaxPercentage
				tTaxTypeDesc=request.Form("TaxDesc"&i)
				tTaxAdding=taxabletotal*(tTaxPercentage/100)
				tTax=cdbl(tTax)+cdbl(tTaxAdding)
				ptaxDetailString=ptaxDetailString&tTaxTypeDesc&"|"&tTaxAdding&","
			else
				tTaxAdding=request.Form("taxTotal"&i)
				if tTaxAdding="" then
				tTaxAdding="0"
				end if
				tTaxTypeDesc=request.Form("TaxDesc"&i)
				tTax=cdbl(tTax)+cdbl(tTaxAdding)
				ptaxDetailString=ptaxDetailString&tTaxTypeDesc&"|"&tTaxAdding&","
			end if
		next
	else
		tTaxPercentage=request.Form("calculateTaxPercentage")
		if tTaxPercentage="" then
		tTaxPercentage="0"
		end if
		if tTaxPercentage<>"" AND isNumeric(tTaxPercentage) AND tTaxPercentage<>"0" then
			session("adminTaxPercentage")=tTaxPercentage
			tTax=taxabletotal*(tTaxPercentage/100)
		else
			tTax=request.Form("taxTotal")
			if tTax<>"" then
				tTax=replacecommaToCal(tTax)
			end if
		end if
	end if
	if tTax="" then
		tTax=0
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Calculate Taxes
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Adjust VAT
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	tVAT=request.Form("VATTotal")
	if tVAT="" then
		tVAT=0
	end if
	tVAT=replacecommaToCal(tVAT)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Adjust VAT
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



	'--------------------
	'calculate everything and UPdate database
	tTotal=subtotal+tTax
	tshippingProvider=request.Form("shippingProvider")
	tshippingService=request.Form("shippingService")
	tshippingServiceCode=request.Form("shippingServiceCode")
	tTotalWeight=request.Form("weightTotal")
	if tShipping=0 and tHandling=0 AND tshippingProvider="" AND tshippingService="" then
		tShippingString="No shipping is required for this order."
	else
		if trim(tshippingProvider)="" then
			tshippingProvider="CUSTOM"
		end if
		if trim(tshippingService)="" then
			tshippingService="Custom"
		end if
		tShippingString=tshippingProvider&","&tshippingService&","&replacecommaToDB(tShipping)&","&replacecommaToDB(tHandling)&",0,"&tshippingServiceCode
	end if
	tpaymentType=request.Form("paymentType")
	tPaymentString=tpaymentType&"||"&replacecommaToDB(tpaymentCharges)
	pcv_PaymentStatus=request("pcv_PaymentStatus")
	if pcv_PaymentStatus="" then
		pcv_PaymentStatus=0
	end if

	'GGG Add-on start
	tTotal=tTotal+subGWTotal
	'GGG Add-on end

	'GGG Add-on start
	'Update Gift Certificate Amount
	query="select pcOrd_GCDetails,pcOrd_GCAmount,pcOrd_GcCode,pcOrd_GcUsed from Orders WHERE idOrder="&qryID&";"
	set rs=connTemp.execute(query)
	pGiftCode=rs("pcOrd_GcCode")
	pGiftUsed=rs("pcOrd_GcUsed")
	if pGiftUsed<>"" then
	else
		pGiftUsed=rs("pcOrd_GCAmount")
		if pGiftUsed<>"" then
		else
			pGiftUsed=0
		end if
	end if
	
	tTotal=tTotal-pGiftUsed
	'GGG Add-on end
	'Update Order Details File
	pDetails=""
	query="SELECT products.description,products.sku,((ProductsOrdered.unitPrice*ProductsOrdered.quantity)) AS IAmount,ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_OptionsArray FROM products INNER JOIN ProductsOrdered ON products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idOrder="&qryID&";"
	set rs=ConnTemp.execute(query)
	Do while not rs.eof
		pcv_PrdName=rs("description")
		pcv_PrdSKU=rs("sku")
		pcv_PrdAmount=rs("IAmount")
		pcv_PrdQty=rs("quantity")
		pcv_strOptionsArray=rs("pcPrdOrd_OptionsArray")
		pDetails = pDetails & "  Amount: ||"& pcv_PrdAmount & " Qty:" &pcv_PrdQty& "  SKU #:" &pcv_PrdSKU & " - " &pcv_PrdName& " " & " " & pcv_strOptionsArray & Vbcrlf

		rs.MoveNext
	Loop
	set rs=nothing
	if pDetails<>"" then
		pDetails = replace(pDetails,"'","''")
		pDetails=replace(pDetails,"''''","''")
	end if

	query="UPDATE orders SET details='" & pDetails & "',pcOrd_PaymentStatus=" & pcv_PaymentStatus & ",total="&replacecommaToDB(tTotal)&",taxAmount="&replacecommaToDB(tTax)&", ord_VAT="&replacecommaToDB(tVAT)&", shipmentDetails='"&replace(tShippingString,"'","''")&"',paymentDetails='"&replace(tPaymentString,"'","''")&"',discountDetails='"&replace(tDiscountString,"'","''")&"',iRewardPoints="&useRewards&", iRewardValue="&iDollarValue&", iRewardPointsCustAccrued="&curRewardPointsAccrued&",SRF=0,taxDetails='"&ptaxDetailString&"',pcOrd_GcUsed=" & pGiftUsed & ",pcOrd_GWTotal=" & subGWTotal & ", pcOrd_CatDiscounts=" & replacecommaToDB(CatDiscTotal) & ", pcOrd_ShipWeight="&tTotalWeight&" WHERE idOrder="&tIdOrder&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	query="SELECT idProductOrdered,idProduct,quantity FROM ProductsOrdered WHERE idOrder="&qryID&";"
	set rs=connTemp.execute(query)

	do while not rs.eof
		pidProductOrdered=rs("idProductOrdered")
		pidProduct=rs("idProduct")
		newqty=rs("quantity")
		mCount=session("admin_cp_edit_count")
		haveit=0
		For j=1 to mCount
			if clng(pidProductOrdered)=clng(session("admin_cp_" & qryID & "_idprd" & j)) then
				haveit=1
				pcv_qty=session("admin_cp_" & qryID & "_qty" & j)
				newqty=clng(newqty)-clng(pcv_qty)
				query="UPDATE Products SET stock=stock-" & newqty & ",sales=sales+" & newqty & " WHERE idproduct=" & pidProduct
				set rs1=server.CreateObject("ADODB.RecordSet")
				set rs1=conntemp.execute(query)
				set rs1=nothing
				exit for
			end if
		Next
		if haveit=0 then
			query="UPDATE Products SET stock=stock-" & newqty & ",sales=sales+" & newqty & " WHERE idproduct=" & pidProduct
			set rs1=server.CreateObject("ADODB.RecordSet")
			set rs1=conntemp.execute(query)
			set rs1=nothing
		end if
		rs.MoveNext
	loop
	set rs=nothing

	'update customer's rewards
	pIdCustomer=request.Form("idCustomer")
	call closedb()
	session("adminCurrentPrice"&tempCnt)=""
	response.redirect "AdminEditOrder.asp?ido="&request("ido")&"&mainUpdate=yes"
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: POST BACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: ON RESTORE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if request.QueryString("action")="restore" then
	dim tempCnt, tempIDO, tempIDP, tempIdprdO
	tempCnt=request.QueryString("Cnt")
	tempIDO=request.QueryString("ido")
	tempIDP=request.QueryString("idprd")
	tempIdprdO=request.QueryString("idprdO")
	call opendb()
	'// Find current price of item
	query="SELECT price FROM products WHERE idProduct="&tempIDP
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	session("adminCurrentPrice"&tempCnt)= rs("price")
	set rs=nothing
	'// Reset our QDiscounts
	query="UPDATE ProductsOrdered SET QDiscounts=" & 0 & " WHERE idProductOrdered="&tempIdprdO&" AND idOrder="&tempIDO&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing
	call closedb()
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: ON RESTORE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
call opendb()
query="SELECT orders.pcOrd_IDEvent,orders.pcOrd_PaymentStatus,orders.pcOrd_CATDiscounts,orders.idOrder, orders.orderDate, orders.idCustomer, orders.shippingStateCode,	orders.shippingCountryCode, orders.shippingZip, pcOrd_ShipWeight,customers.name, customers.lastname, customers.customerType,orders.stateCode, orders.zip, orders.CountryCode, orders.shippingAddress FROM orders, customers WHERE orders.idCustomer=customers.idCustomer AND orders.idOrder="&qryID&";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
dim wrngID
wrngID=0
if rs.eof then
	'no orders match this id
	wrngID=1
	call closedb()
else
	'GGG Add-on start
	gIDEvent=rs("pcOrd_IDEvent")
	if gIDEvent<>"" then
	else
	gIDEvent="0"
	end if
	'GGG Add-on end

	'Start SDBA
	pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
	if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
		pcv_PaymentStatus=0
	end if
	'End SDBA
	CatDiscounts= rs("pcOrd_CatDiscounts")
	orderDate=rs("orderDate")
	pIdCustomer=rs("idCustomer")
	pshippingStateCode=rs("shippingStateCode")
	pshippingCountryCode=rs("shippingCountryCode")
	pshippingZip=rs("shippingZip")
	pOrderShipWeight=rs("pcOrd_ShipWeight")
	fname=rs("name")
	lname=rs("lastName")
	pcustomerType=rs("customerType")
	Session("customerType")=pcustomerType

	pCustSateCode=rs("stateCode")
	pCustZip=rs("zip")
	pCustCountryCode=rs("CountryCode")
	pshippingAddress=rs("shippingAddress")

	if pshippingAddress="" OR IsNull(pshippingAddress) then
		pshippingStateCode=pCustSateCode
		pshippingCountryCode=pCustCountryCode
		pshippingZip=pCustZip
	end if

end if

'// Check if the Customer is European Union
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<script>
	function winSale(fileName)
	{
		myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=650,height=300')
		myFloater.location.href=fileName;
	}
</script>
<form action="AdminEditOrder.asp" method="post" name="EditOrder" class="pcForms" onsubmit="return validate()">
	<input name="idcustomer" type="hidden" id="idcustomer" value="<%=pIdCustomer%>">
	<% if wrngID=0 then %>
		<table class="pcCPcontent">
			<tr>
				<td><h2>You are editing order #<%=int(qryID)+scpre%> | <a href="ordDetails.asp?id=<%=qryID%>">Back to Order Details &gt;&gt;</a></h2></td>
			</tr>
	  <tr>
		<td>Order Date: <%=ShowDateFrmt(orderDate)%> | Customer Name: <a href="modCusta.asp?idcustomer=<%=pIdCustomer%>"><%=fname&" "&lname%></a></td>
			</tr>
			<tr>
				<td class="pcCPspacer" align="center">
					<% If (request.QueryString("action")="upds") AND (request.QueryString("removed")<>"yes") then %>
						<div class="pcCPmessageSuccess">
							<p>You have successfully added a new product to this order.</p>
							  <p>&nbsp;</p>
							  <p style="font-weight: normal;">NOTE: shipping charges are not recalculated automatically. Use the &quot;Check Real-Time Rates&quot; to obtain updated shipping rates.</p>
							  <p>&nbsp;</p>
							  <p style="font-weight: normal;">NOTE: promotion-related discounts are NOT applied automatically if you change the order quantity</p>
							  <p>&nbsp;</p>
							  <p style="font-weight: normal;">Review the order, change any settings that might need to be changed, and then click on &quot;Update Order&quot; at the bottom of the page to fully update the order details.</p>
						</div>
					<%else
						If request.QueryString("removed")="yes" then%>
						<div class="pcCPmessageSuccess">
							<p>You have successfully removed product(s) from this order.</p>
						  <p>&nbsp;</p>
						  <p style="font-weight: normal;">NOTE: shipping charges are not recalculated automatically. Use the &quot;Check Real-Time Rates&quot; to obtain updated shipping rates.</p>
						  <p>&nbsp;</p>
						  <p style="font-weight: normal;">NOTE: promotion-related discounts are NOT applied automatically if you change the order quantity</p>
						  <p>&nbsp;</p>
						  <p style="font-weight: normal;">Review the order, change any settings that might need to be changed, and then click on &quot;Update Order&quot; at the bottom of the page to fully update the order details.</p>
						 </div>
						<%end if%>
					<% end if %>
		  <% if request.QueryString("mainUpdate")="yes" then %>
						<div class="pcCPmessageSuccess">
							<p>The order has been updated.</p>
							  <p>&nbsp;</p>
							  <p style="font-weight: normal;">NOTE: when you update an order, <u>shipping charges</u> are not recalculated automatically.</p>
							  <p>&nbsp;</p>
							  <p style="font-weight: normal;">NOTE: promotion-related discounts are NOT applied automatically if you change the order quantity</p>
							  <p>&nbsp;</p>
							  <p style="font-weight: normal;">If you have note already done so, use the &quot;Check Real-Time Rates&quot; to obtain updated shipping rates, and then click on &quot;Update Order&quot; at the bottom of the page to save the changes to the database.</p>
						</div>
		  <% end if %>
				</td>
			</tr>
		</table>

		<table class="pcCPcontent">
			<tr>
				<td>
					<table class="pcCPcontent">
						<tr>
							<th>SKU</th>
							<th>Product</th>
							<th>Qty</th>
							<th>Unit Price&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=312')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
							<th colspan="2">Price&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=312')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
						</tr>

						<%
						Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
						Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
						Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions

						call opendb()
						query="SELECT ProductsOrdered.pcSC_ID,ProductsOrdered.pcPO_GWOpt, ProductsOrdered.pcPO_GWNote, ProductsOrdered.pcPO_GWPrice, ProductsOrdered.idProductOrdered,  ProductsOrdered.idOrder, ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.unitPrice, ProductsOrdered.xfdetails, ProductsOrdered.QDiscounts, ProductsOrdered.ItemsDiscounts"
						'BTO ADDON-S
						if scBTO=1 then
							query=query&", ProductsOrdered.idconfigSession"
						end if
						'BTO ADDON-E
						query=query&", products.description, products.weight, products.sku, products.notax,products.stock, products.nostock, products.pcProd_BackOrder, pcprod_QtyToPound FROM products INNER JOIN ProductsOrdered ON products.idProduct = ProductsOrdered.idProduct WHERE (((ProductsOrdered.idOrder)="&qryID&"));"

						set rs=server.CreateObject("ADODB.RecordSet")
						rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText

						session("admin_cp_edit_idorder")=qryID
						session("admin_cp_edit_count")=0

						Dim pcv_inTotalCount
						if rs.eof then
							pcv_inTotalCount = 0
						else
							pcv_inTotalCount = rs.recordcount
						end if

						pCnt=0
						pweightTotal=0
						pTotalCartQty=0
						do until rs.eof
							pCnt=pCnt+1
							session("admin_cp_edit_count")=pCnt
							pcSCID=rs("pcSC_ID")
							if pcSCID="" OR (IsNull(pcSCID)) then
								pcSCID=0
							end if
							'GGG Add-on start
							pcv_GWOpt=rs("pcPO_GWOpt")
							if pcv_GWOpt<>"" then
							else
								pcv_GWOpt="0"
							end if
							if pcv_GWOpt<>"0" then
								query="select pcGW_OptName from pcGWOptions where pcGW_IDOpt=" & pcv_GWOpt
								set rsGW=connTemp.execute(query)
								pcv_GWDesc=rsGW("pcGW_OptName")
							end if
							pcv_GWNote=rs("pcPO_GWNote")
							pcv_GWPrice=rs("pcPO_GWPrice")
							if (pcv_GWOpt<>"") and (pcv_GWOpt<>"0") then
								pcv_HaveGWOpt=1
							else
								pcv_HaveGWOpt=0
							end if

							if pcv_HaveGWOpt=0 then
								query="select pcGW_IDOpt from pcGWOptions"
								set rsGW=conntemp.execute(query)
								if not rsGW.eof then
									query="select pcPE_IDProduct from pcProductsExc where pcPE_IDProduct=" & rs("idProductOrdered")
									set rsGW=connTemp.execute(query)
									if rsGW.eof then
										pcv_HaveGWOpt=1
									end if
								end if
							end if
							'GGG Add-on end

							pIdProductOrdered=rs("idProductOrdered")
							pIdOrder=rs("idOrder")
							pIdProduct=rs("idProduct")
							pQuantityt=rs("quantity")

							session("admin_cp_" & qryID & "_idprd" & pCnt)=pIdProductOrdered
							session("admin_cp_" & qryID & "_id" & pCnt)=pIdProduct
							session("admin_cp_" & qryID & "_qty" & pCnt)=pQuantityt

							'// Product Options Arrays
							pcv_strSelectedOptions = ""
							pcv_strOptionsPriceArray = ""
							pcv_strOptionsArray = ""
							pcv_strSelectedOptions = rs("pcPrdOrd_SelectedOptions") ' Column 11
							pcv_strOptionsPriceArray = rs("pcPrdOrd_OptionsPriceArray") ' Column 25
							pcv_strOptionsArray = rs("pcPrdOrd_OptionsArray") ' Column 4

							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' START: Get the total Price of all options
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							pOpPrices=0
							dim pcv_tmpOptionLoopCounter, pcArray_TmpCounter
							If len(pcv_strOptionsPriceArray)>0 then

								pcArray_TmpCounter = split(pcv_strOptionsPriceArray,chr(124))
								For pcv_tmpOptionLoopCounter = 0 to ubound(pcArray_TmpCounter)
									pOpPrices = pOpPrices + pcArray_TmpCounter(pcv_tmpOptionLoopCounter)
								Next

							end if

							if NOT isNumeric(pOpPrices) then
								pOpPrices=0
							end if

							'// Apply Discounts to Options Total
							'   >>> call function "pcf_DiscountedOptions(OriginalOptionsTotal, ProductID, Quantity, CustomerType)" from stringfunctions.asp
							Dim pcv_intDiscountPerUnit
							if request.QueryString("action")<>"restore" then
								pOpPrices = pcf_DiscountedOptions(pOpPrices, pIdProduct, pQuantityt, pcustomerType)
							end if
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' END: Get the total Price of all options
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

							pOnlinePrice=rs("unitPrice")
							'// Adjust the unit price "online price" for options
							pOnlinePrice=(pOnlinePrice-pOpPrices)

							pxfdetails=rs("xfdetails")
							pQDiscounts=rs("QDiscounts")
							if pQDiscounts<>"" then
							else
								pQDiscounts="0"
							end if
							pItemsDiscounts=rs("ItemsDiscounts")
							if pItemsDiscounts<>"" then
							else
								pItemsDiscounts="0"
							end if
							if scBTO=1 then
								pIdConfigSession=rs("idconfigSession")
							end if
							if NOT request.QueryString("action")="restore" then
								if session("adminCurrentPrice"&pCnt)<>"" then
									session("adminCurrentPrice"&pCnt)= request.form("currentPrice")
								else
									session("adminCurrentPrice"&pCnt)= pOnlinePrice
								end if
							else
								if int(tempCnt)<>int(pCnt) then
									session("adminCurrentPrice"&pCnt)= pOnlinePrice
								else
									session("adminCurrentPrice"&pCnt)= session("adminCurrentPrice"&pCnt)
								end if
							end if

							pProductName=rs("description")
							pWeight=rs("weight")

							pSKU=rs("sku")
							pNoTax=rs("notax")

							session("adminpQTY"&pCnt)=pQuantityt
							pTotalCartQty=pTotalCartQty+pQuantityt
							if session("adminpQTY"&pCnt)="" then
								session("adminpQTY"&pCnt)="1"
							end if
							'Start SDBA
							pcv_stock=rs("stock")
							if IsNull(pcv_stock) or pcv_stock="" then
								pcv_stock=0
							end if
							if clng(pcv_stock)<0 then
								pcv_stock=0
							end if
							pcv_nostock=rs("nostock")
							if IsNull(pcv_nostock) or pcv_nostock="" then
								pcv_nostock=0
							end if
							pcv_backorder=rs("pcProd_backorder")
							if IsNull(pcv_backorder) or pcv_backorder="" then
								pcv_backorder=0
							end if
							tmp_stockmsg=""
							if pcv_stock>0 then
								tmp_stockmsg="In-Stock"
							else
								if pcv_nostock<>0 then
									tmp_stockmsg="Not Available - Disregard Stock"
								else
									if pcv_backorder=1 then
										tmp_stockmsg="Back-Order"
									else
										tmp_stockmsg="Not Available"
									end if
								end if
							end if
							pcv_QtyToPound=rs("pcprod_QtyToPound")
							if pcv_QtyToPound>0 then
								pWeight=(16/pcv_QtyToPound)
								if scShipFromWeightUnit="KGS" then
									pWeight=(1000/pcv_QtyToPound)
								end if
							end if
							pweightTotal=pweightTotal+(pWeight*pQuantityt)

							'End SDBA%>
							<tr>
								<td><%=pSKU%></td>
								<td><a href="FindProductType.asp?idproduct=<%=pIdProduct%>" target="_blank"><%=pProductName%></a>
								<%if pcSCID>"0" then
									query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon FROM pcSales_Completed WHERE pcSales_Completed.pcSC_ID=" & pcSCID & ";"
									set rsS=Server.CreateObject("ADODB.Recordset")
									set rsS=conntemp.execute(query)

									if not rsS.eof then
										pcSCID=rsS("pcSC_ID")
										pcSCName=rsS("pcSC_SaveName")
										pcSCIcon=rsS("pcSC_SaveIcon") %>

										&nbsp;<a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="../pc/catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>" style="vertical-align: middle"></a>
									<%end if
									set rsS=nothing
								end if%>
								<br>
								<i>(<%=tmp_stockmsg%>)</i></td>
								<td><input name="QTY<%=pCnt%>" type="text" id="qty<%=pCnt%>" value="<%=session("adminpQTY"&pCnt)%>" size="3"></td>
								<td>
								<input name="optionsLineTotal<%=pCnt%>" type="hidden" value="<%=pOpPrices%>">
								<input name="strOptionsPriceArray<%=pCnt%>" type="hidden" value="<%=pcv_strOptionsPriceArray%>">
								<%
								'// Original Price for ReCalulating Quantity Discounts
								OriginalPrice = (session("adminCurrentPrice"&pCnt) + pOpPrices)
								%>
								<input name="OriginalPrice<%=pCnt%>" type="hidden" value="<%=OriginalPrice%>">
								<input name="SavedUnitPrice<%=pCnt%>" type="hidden" value="<%=money(session("adminCurrentPrice"&pCnt))%>">
								<input name="unitPrice<%=pCnt%>" type="text" id="unitPrice<%=pCnt%>" value="<%=money(session("adminCurrentPrice"&pCnt))%>" size="10">&nbsp;
								<%'GGG Add-on
								if gIDEvent="0" then%><a href="AdminEditOrder.asp?action=restore&Cnt=<%=pCnt%>&ido=<%=pIdOrder%>&idprd=<%=pIdProduct%>&idprdO=<%=pIdProductOrdered%>" title="Restore unit price to current online price for this product">Restore</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=313')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a><%end if 'GGG Add-on end%>
								&nbsp;&nbsp;
								<input name="idorder<%=pCnt%>" type="hidden" value="<%=pIdOrder%>">
								<input name="idProduct<%=pCnt%>" type="hidden" value="<%=pIdProduct%>">
								<input name="idProductOrdered<%=pCnt%>" type="hidden" value="<%=pIdProductOrdered%>">
								</td>

								<%
								'BTO Additional Charges
								Charges=0
								if scBTO=1 then
									if pIdConfigSession<>"0" then
										query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
										set rsConfigObj=conntemp.execute(query)
										if err.number <> 0 then
											response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails: "&err.description)
										end if
										stringCProducts=rsConfigObj("stringCProducts")
										stringCValues=rsConfigObj("stringCValues")
										stringCCategories=rsConfigObj("stringCCategories")
										ArrCProduct=Split(stringCProducts, ",")
										ArrCValue=Split(stringCValues, ",")
										ArrCCategory=Split(stringCCategories, ",")
										set rsConfigObj=nothing
										if ArrCProduct(0)<>"na" then
											for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
												query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))"
												set rsConfigObj=connTemp.execute(query)
												if (CDbl(ArrCValue(i))>0)then
													Charges=Charges+cdbl(ArrCValue(i))
												end if
												set rsConfigObj=nothing
											next

										end if
									end if
								end if 'BTO Additional Charges

								extprice=int(pQuantityt)*cdbl(session("adminCurrentPrice"&pCnt))-cdbl(pItemsDiscounts)+cdbl(Charges)%>
								<td><%=scCurSign&money(extprice)%></td>
								<td nowrap="nowrap">
								<div align="right">
								<% if pcv_IsEUMemberState=2 then %>
									<%
									if pNoTax=0 then
									%>
										<input name="taxProduct<%=pCnt%>" type="checkbox" id="taxProduct<%=pCnt%>" value="YES" checked class="clearBorder" />
									<% else %>
										<input name="taxProduct<%=pCnt%>" type="checkbox" id="taxProduct<%=pCnt%>" value="YES" class="clearBorder" />
									<% end if %>
									Tax&nbsp;
								<% end if %>
								<% if pcv_inTotalCount>1 then %>
									<a href="javascript:if (confirm('You are about to delete this Product from this order. This action can not be undone, are you sure you want to complete this action?')) location='AdminEditOrder.asp?del=YES&action=upd&delPrd=<%=pIdProductOrdered%>&ido=<%=pIdOrder%>'"><img src="images/delete2.gif" width="23" height="18" border="0"></a>
								<% end if %>
								</div>
								</td>
							</tr>

							<% 'BTO ADDON-S
							if scBTO=1 then
								if pIdConfigSession<>"0" then
									query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
									set rsConfigObj=connTemp.execute(query)
									stringProducts=rsConfigObj("stringProducts")
									stringValues=rsConfigObj("stringValues")
									stringCategories=rsConfigObj("stringCategories")
									stringQuantity=rsConfigObj("stringQuantity")
									stringPrice=rsConfigObj("stringPrice")
									ArrProduct=Split(stringProducts, ",")
									ArrValue=Split(stringValues, ",")
									ArrCategory=Split(stringCategories, ",")
									ArrQuantity=Split(stringQuantity, ",")
									ArrPrice=Split(stringPrice, ",")
									%>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3">
											<table class="pcCPcontent">
												<tr>
													<td colspan="2">
														Customizations
														<%'GGG Add-on start
														if gIDEvent="0" then%>
															&nbsp;-&nbsp;<a href="bto_Reconfigure.asp?idp=<%=pIdProductOrdered%>&ido=<%=qryid%>">Edit Configuration</a>
														<%end if 'GGG Add-on end%>
													</td>
												</tr>

												<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
													query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
													set rsConfigObj=connTemp.execute(query)
													query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and SpecProduct=" & pIdProduct
													set rsObj1=conntemp.execute(query)
													if rsObj1.eof then
														btDisplayQF=False
													else
														btDisplayQF=rsObj1("displayQF")
													end if
													strCategoryDesc=rsConfigObj("categoryDesc")
													strDescription=rsConfigObj("description")
													strSKU=rsConfigObj("sku")
													%>
													<tr valign="top" class="pcSmallText">
														<td width="29%"><%=strCategoryDesc%>:</td>
														<td width="71%"><%=strSKU%> - <%=strDescription%><%if btDisplayQF=True then%>
																- QTY: <%=ArrQuantity(i)%><%end if%></td>
													</tr>
													<%
													set rsConfigObj=nothing
												next
												set rsConfigObj=nothing
												%>
											</table>
										</td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
									</tr>
								<% end if
							end if %>
							<% query="select * from discountsPerQuantity WHERE idproduct="&pIdProduct&";"
							set rstemp=server.CreateObject("ADODB.RecordSet")
							set rstemp=connTemp.execute(query)
							if NOT rstemp.eof then %>
								<tr bgcolor="#FFFFFF">
									<td>&nbsp;</td>
									<td><a href="javascript:win2('../pc/priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=pidProduct%>')">View Quantity Discounts</a></td>
									<td></td>
									<td>QUANTITY DISCOUNTS:</td>
									<td colspan="2" align="left"><%=scCurSign&"-"&money(pQDiscounts)%></td>
								</tr>
							<% end if %>
							<!-- start options -->
							<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' START: SHOW PRODUCT OPTIONS
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// CHECK FOR OPTIONS
							' SELECT DATA SET
							' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
							query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.pcProdOpt_ID, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
							query = query & "FROM products "
							query = query & "INNER JOIN ( "
							query = query & "pcProductsOptions INNER JOIN ( "
							query = query & "optionsgroups "
							query = query & "INNER JOIN options_optionsGroups "
							query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
							query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
							query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
							query = query & "WHERE products.idProduct=" & pidProduct &" "
							query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
							query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order;"
							set rsCheckOptions=server.createobject("adodb.recordset")
							set rsCheckOptions=conntemp.execute(query)
							if err.number<>0 then

							end if

							If NOT rsCheckOptions.eof Then
								pcv_intOptionsExist = 1
							Else
								pcv_intOptionsExist = 2
							End If
							set rsCheckOptions = nothing

							if (len(pcv_strSelectedOptions)>0) AND (pcv_strSelectedOptions<>"NULL") then
								%>
								<tr bgcolor="#FFFFFF">
									<td>&nbsp;</td>
									<td><b>Product Options:</b></td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								</tr>
								<%
								'#####################
								' START LOOP
								'#####################

								'// Generate Our Local Arrays from our Stored Arrays

								' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers
								pcArray_strSelectedOptions = ""
								pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))

								' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
								pcArray_strOptionsPrice = ""
								pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))

								' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
								pcArray_strOptions = ""
								pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))

								' Get Our Loop Size
								pcv_intOptionLoopSize = 0
								pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)

								' Start in Position One
								pcv_intOptionLoopCounter = 0

								' Display Our Options
								For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
									%>
									<tr valign="top">
										<td>&nbsp;</td>
										<td>
										<%=pcArray_strOptions(pcv_intOptionLoopCounter) %>
										</td>
										<td>&nbsp;</td>
										<%
										'// Saved Options Price
										tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
										if tempPrice="" or tempPrice=0 then
											response.write "<td colspan=2>&nbsp;</td>"
										else
											'// Adjust for Quantity Discounts
											tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice)
											%>
											<td align="left">
												<%=scCurSign&money(tempPrice)%>
											</td>
											<td>
												<%
												tAprice=(tempPrice*Cdbl(session("adminpQTY"&pCnt)))
												response.write scCurSign&money(tAprice)
												%>
											</td>

										<% end if %>
									</tr>
								<% Next
								'#####################
								' END LOOP
								'#####################

								If pcv_intOptionsExist = 1 Then
									%>
									<tr>
										<td>&nbsp;</td>
										<td colspan="5">
										<% '// GGG Add-on start
										if gIDEvent="0" then %>
											<a href="javascript:;" onClick="newWindow2('options_popup.asp?idProductOrdered=<%=pIdProductOrdered%>&idOptionArray=<%=pcv_strSelectedOptions%>&idproduct=<%=pidProduct%>','window2')">Edit Option(s)</a>
										<% end if
										'// GGG Add-on end %>
										</td>
									</tr>
								<% end if
							else
								If pcv_intOptionsExist = 1 Then %>
									<tr bgcolor="#FFFFFF">
										<td>&nbsp;</td>
										<td colspan="5">
										<%
										'// GGG Add-on start
										if gIDEvent="0" then
										%>
										<a href="javascript:;" onClick="newWindow2('options_popup.asp?idProductOrdered=<%=pIdProductOrdered%>&idOptionArray=<%=pcv_strSelectedOptions%>&idproduct=<%=pidProduct%>','window2')">
											Add Option(s)
										</a>
										<%
										end if
										'// GGG Add-on end
										%>
										</td>
									</tr>
								<% end if
							End If
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' END: SHOW PRODUCT OPTIONS

							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							%>
							<!-- end options -->
							<% 'BTO ADDON-S
							if scBTO=1 then
								if pIdConfigSession<>"0" then
									query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
									set rsConfigObj=connTemp.execute(query)
									stringCProducts=rsConfigObj("stringCProducts")
									stringCValues=rsConfigObj("stringCValues")
									stringCCategories=rsConfigObj("stringCCategories")
									ArrCProduct=Split(stringCProducts, ",")
									ArrCValue=Split(stringCValues, ",")
									ArrCCategory=Split(stringCCategories, ",")
									if ArrCProduct(0)<>"na" then%>
										<tr bgcolor="#FFFFFF">
											<td>&nbsp;</td>
											<td colspan="3">
												<table width="100%" border="0" cellspacing="0" cellpadding="2">
													<tr class="small">
														<td colspan="2">ADDITIONAL CHARGES</td>
													</tr>
													<% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
														query="SELECT categories.categoryDesc, products.description, products.sku FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))"
														set rsConfigObj=connTemp.execute(query)%>
														<tr valign="top" class="small">
														<td width="29%"><%=rsConfigObj("categoryDesc")%>:</td>
														<td width="71%"><%=rsConfigObj("sku")%> - <%=rsConfigObj("description")%></td>
														</tr>
														<% set rsConfigObj=nothing
													next
													set rsConfigObj=nothing %>
													<tr valign="top" class="small">
													  <td>&nbsp;</td>
													  <td><div align="right">
													  <%'GGG Add-on
													  if gIDEvent="0" then%>
													  <a href="bto_Reconfigure.asp?idp=<%=pIdProductOrdered%>&ido=<%=qryid%>">Edit Configuration</a>
													  <%end if 'GGG Add-on end%>
													  </div></td>
													</tr>
												</table>
											</td>
											<td>&nbsp;</td>
											<td>&nbsp;</td>
										</tr>
									<% end if 'Have Additional Charges
								end if
							end if

							'GGG Add-on Start
							if pcv_HaveGWOpt=1 then%>
								<tr bgcolor="#FFFFFF">
									<td>&nbsp;</td>
									<td colspan="3">
									<%if (pcv_GWOpt<>"") and (pcv_GWOpt<>"0") then%>
										<b>Gift Wrapping:</b> <%=pcv_GWDesc%>
										<br />
										<b>Gift Note:</b> <%=pcv_GWNote%>
										<br />
									<%end if%>
									<a href="javascript:;" onClick="newWindow2('ggg_customGW_popup.asp?idProductOrdered=<%=pIdProductOrdered%>&idOrder=<%=qryID%>&IDGW=<%=pcv_GWOpt%>','window2')"><%if (pcv_GWOpt<>"") and (pcv_GWOpt<>"0") then%>Edit Wrapping Option<%else%>Add Wrapping Option<%end if%></a></td>
									<td><div align="right"><%if (pcv_GWOpt<>"") and (pcv_GWOpt<>"0") then%><%=scCurSign & money((pcv_GWPrice*session("adminpQTY"&pCnt)))%><%end if%></div></td>
									<td width="24"><% if intOrderStatus<3 then %><div align="right"><%if (pcv_GWOpt<>"") and (pcv_GWOpt<>"0") then%><a href="javascript:if (confirm('You are about to delete this gift wrapping option from this product. This action can not be undone, are you sure you want to complete this action?')) newWindow2('ggg_customGW_popup.asp?idProductOrdered=<%=pIdProductOrdered%>&idOrder=<%=qryID%>&action=del&IDGW=<%=pcv_GWOpt%>','window2');"><img src="images/delete2.gif" width="23" height="18" border="0"></a><%end if%><% end if %></div>
									</td>
								</tr>
							<%end if
							'GGG Add-on end %>
							<% if pxfdetails<>"" then
								xfArray=split(pxfdetails,"|")
								for xf=0 to ubound(xfArray)
									tempXf=xfArray(xf)
									if trim(tempXf)<>"" then
									if InStr(tempXf,":")>0 then
									xSplitArray=split(tempXf,":")
									%>
									<tr bgcolor="#FFFFFF">
										<td>&nbsp;</td>
										<td><%=xSplitArray(0)%>:&nbsp;<%=xSplitArray(1)%>&nbsp;&nbsp;<a href="javascript:;" onClick="newWindow2('customInput_popup.asp?idProductOrdered=<%=pIdProductOrdered%>&c=<%=xf%>&idOrder=<%=qryID%>','window2')">Edit</a></td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										<td><div align="right"></div></td>
										<td width="24"><div align="right"></div>
										</td>
									</tr>
								<% end if
								end if
								next
							end if %>
							<% rs.movenext
						loop %>
						<input name="tCnt" type="hidden" value="<%=pCnt%>" size="4">
					</table>
				</td>
			</tr>
			<tr>
				<td><hr></td>
			</tr>
			<tr>
				<td>
					<%'GGG Add-on start
					if gIDEvent="0" then%>
						<p><a href="LocateProduct.asp?ido=<%=qryID%>">Add Product to Order</a></p>
					<%else%>
						<p><a href="ggg_AdminViewGR.asp?ido=<%=qryID%>&IDEvent=<%=gIDEvent%>">Add Product to Order</a></p>
					<%end if
					'GGG Add-on end%>
				</td>
			</tr>
		</table>

		<% query="SELECT orderDate, idCustomer, total, taxAmount, ord_VAT, shipmentDetails, paymentDetails, discountDetails, iRewardPoints,iRewardValue,iRewardPointsCustAccrued,taxDetails,pcOrd_GcCode,pcOrd_GcUsed FROM orders WHERE idOrder="&qryID&";"

		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		pOrderDate=rs("orderDate")
		pIdCustomer=rs("idCustomer")
		ptotal=rs("total")
		ptaxAmount=rs("taxAmount")
		pord_VAT=rs("ord_VAT")
		shipmentArray=rs("shipmentDetails")
		paymentArray=rs("paymentDetails")
		discountArray=rs("discountDetails")
		pIRewardPoints=rs("iRewardPoints")
		pIRewardValue=rs("iRewardValue")
		totalOrderAccrued=rs("iRewardPointsCustAccrued")
		ptaxDetails=rs("taxDetails")
		'GGG Add-on start
		pGiftCode=rs("pcOrd_GcCode")
		pGiftUsed=rs("pcOrd_GcUsed")
		if pGiftUsed<>"" then
		else
			pGiftUsed=0
		end if
		'GGG Add-on end
		shipSplit=split(shipmentArray,",")
		varShip="1"
		if ubound(shipSplit)>1 then
			if NOT isNumeric(trim(shipSplit(2))) then
				varShip="0"
				shipPrice=0
				shipHandling=0
			else
				shipProvider=shipSplit(0)
				shipService=shipSplit(1)
				shipPrice=trim(shipSplit(2))
				shipPrice=replace(shipPrice,CHR(13),"")
				if ubound(shipSplit)=>3 then
					shipHandling=trim(shipSplit(3))
					if NOT isNumeric(shipHandling) then
						shipHandling=0
					end if
				else
					shipHandling=0
				end if
				if ubound(shipSplit)>4 then
					shipServiceCode=trim(shipSplit(5))
				end if
			end if
		else
			varShip="0"
			shipPrice=0
			shipHandling=0
		end if
		paymentSplit=split(paymentArray,"||")
		paymentType=paymentSplit(0)
		if ubound(paymentSplit)>0 then
			paymentPrice=trim(paymentSplit(1))
			if paymentPrice="" then
				paymentPrice=0
			end if
		else
			paymentPrice=0
		end if
		%>
		<% '// Start Reward Points
		if RewardsActive=1 then %>
			<table class="pcCPcontent">
				<% query="SELECT iRewardPointsAccrued, iRewardPointsUsed FROM customers WHERE idCustomer="&pIdCustomer&";"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				totalCustAccrued=rs("iRewardPointsAccrued")
				totalCustUsed=rs("iRewardPointsUsed")
				if totalCustAccrued="" then
					totalCustAccrued=0
				end if
				if totalCustUsed="" then
					totalCustUsed=0
				end if
				CurrentPoints=totalCustAccrued-totalCustUsed
				tempValue=int(CurrentPoints)*(RewardsPercent/100)
				set rs=nothing

				'if no points were used --
				if pIRewardPoints=0 then %>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<th colspan="2"><%=RewardsLabel%></th>
					</tr>
					<tr>
						<td colspan="2">This customer's current balance is <%=CurrentPoints&" "&RewardsLabel%>. This translates into <%=scCurSign&money(tempValue)%>.
						<input name="totalCustAccrued" type="hidden" id="totalCustAccrued" value="<%=totalCustAccrued%>">
						<input name="totalCustUsed" type="hidden" id="totalCustUsed" value="<%=totalCustUsed%>">
						</td>
					</tr>
					<% '// Cannot apply Reward Points (RP) if customer's balance is 0 or negative
					if CurrentPoints<1 then %>
						<tr>
							<td>You cannot apply <%=RewardsLabel%> against this order if the customer's balance is 0 or negative. You can <a href="modcusta.asp?idcustomer=<%=pIdCustomer%>">adjust the balance</a> and then come back to apply <%=RewardsLabel%> against the order.</td>
						</tr>
					<% else %>
						<tr>
							<td>
							Apply <input name="iRewardPoints" type="text" id="iRewardPoints" size="4"> <%=RewardsLabel%> towards this order.
							<input name="curUsedPointsOrder" type="hidden" id="curUsedPointsOrder" value="0">
							</td>
							<td>&nbsp;</td>
						</tr>
					<% end if '// End cannot apply RP if balance is < 1
				else  %>
					<tr>
						<th colspan="2"><%=RewardsLabel%></th>
					</tr>
					<tr>
					  <td>
						<input name="iRewardPoints" type="text" id="iRewardPoints" value="<%=pIRewardPoints %>" size="4">&nbsp;
						<%=RewardsLabel%> have been applied towards this order.
						<input name="curUsedPointsOrder" type="hidden" id="curUsedPointsOrder" value="<%=pIRewardPoints %>">
									</td>
						<td width="18%"><div align="right"><%=scCurSign&money(pIRewardValue)%></div>
					</td>
				</tr>
				<tr>
					<td colspan="2">This customer's current balance is <%=CurrentPoints&" "&RewardsLabel%>. This translates into <%=scCurSign&money(tempValue)%>.
					<input name="totalCustAccrued" type="hidden" id="totalCustAccrued" value="<%=totalCustAccrued%>">
					<input name="totalCustUsed" type="hidden" id="totalCustUsed" value="<%=totalCustUsed%>">
					</td>
				</tr>
			<% end if %>
			<tr>
				<td colspan="2"><hr></td>
			</tr>
			<tr>
			  <td colspan="2">Customer <% if intOrderStatus > 2 then %>accrued<% else %>will accrue<%end if%>&nbsp;<%=totalOrderAccrued&" "&RewardsLabel%>&nbsp;on this order.</td>
			</tr>
			<tr>
				<td colspan="2">Adjust <%=RewardsLabel%>:
				<input name="iRewardPointsCustAccrued" type="text" id="iRewardPointsCustAccrued" size="4">
				-/+
				<input name="curRewardPointsAccruedOrder" type="hidden" id="curRewardPointsAccruedOrder" value="<%=totalOrderAccrued%>"></td>
			</tr>
			<tr>
				<td colspan="2"><hr></td>
			</tr>
			<tr>
				<td colspan="2"><%=RewardsLabel%> are accrued by a customer when an order is processed. If the order is pending, the customer's balance will not change when the <%=RewardsLabel%> accrued change on this page.</td>
			</tr>
		</table>
	<% end if %>
	<br />
	<table class="pcCPcontent">
		<tr>
			<th colspan="4">Discounts</th>
		</tr>
		<%
		intArryCnt=-1
		discountTotalPrice=0
		if discountArray<>"" AND instr(discountArray, "||") then %>
			<tr>
				<td colspan="4" bgcolor="#FFFFFF">
				This order is currently using the following discounts. Click on <u>Remove</u> to delete a discount.
				Make sure you click on the <u>Update Order</u> button at the bottom of the page to save this change.
				Before the order is saved, if you change your mind you can click on <u>Restore</u> to restore the field values for the discount.
				</td>
			</tr>
			<%
			if instr(discountArray,",") then
				DiscountDetailsArry=split(discountArray,",")
				intArryCnt=ubound(DiscountDetailsArry)
			else
				intArryCnt=0
			end if

			strDiscountTableRow=""
			for k=0 to intArryCnt
				if intArryCnt=0 then
					pTempDiscountDetails=discountArray
				else
					pTempDiscountDetails=DiscountDetailsArry(k)
				end if
				discountPrice = 0
				if instr(pTempDiscountDetails,"- ||") then
					discounts = split(pTempDiscountDetails,"- ||")
					discountType = discounts(0)
					discountPrice = discounts(1)
					discountTotalPrice=discountTotalPrice+discountPrice
				end if

				if discountPrice<>0 then
					%>
					<tr>
						<td nowrap>Description: <input name="DiscountType<%=k%>" type="text" value="<%=DiscountType%>" size="30"></td>
						<td>Amount: <input name="DiscountPrice<%=k%>" type="text" value="<%=money(DiscountPrice)%>" size="10"></td>
						<td colspan="2" align="right">
						<a href="javascript:;" onClick="if(confirm('Are you sure you want to remove this discount from the order?')){ document.EditOrder.DiscountType<%=k%>.value ='';document.EditOrder.DiscountPrice<%=k%>.value ='';document.EditOrder.DiscountType<%=k%>.focus();}">Remove</a> -
						<a href="javascript:;" onClick="document.EditOrder.DiscountType<%=k%>.value ='<%=DiscountType%>';document.EditOrder.DiscountPrice<%=k%>.value ='<%=money(DiscountPrice)%>';document.EditOrder.DiscountType<%=k%>.focus();">Restore</a>
						</td>
					</tr>
				<% end if
			next  %>
			<tr>
				<td colspan="4"><hr></td>
			</tr>
		<%  end if '// if discountArray<>"" AND instr(discountArray, "||") then %>
		<input type="hidden" name="intArryCnt" value="<%=intArryCnt%>">
		<tr>
			<td colspan="4" bgcolor="#FFFFFF">Use the fields below to add a discount.
				The percent field will calculate a discount on that
				percentage of the order. The other field will apply
				a flat discount amount to the order. Enter a description in the <em>Name </em>field.</td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" colspan="4">
			Name:
			<input name="DiscountTypeM" type="text" value="">
			&nbsp;&nbsp;
			Calculate as:
			<input name="calculateDisPercentageM" id="calculateDisPercentageM" type="text" value="<%=session("adminDisPercentage")%>" size="4">&nbsp;% of total order <u>or</u> flat amount: <input name="discountsM" id="discountsM" type="text" value="" size="10">
			<input name="ido" type="hidden" value="<%=qryID%>">
			<script language="javascript">
				document.getElementById('calculateDisPercentageM').value='';
				document.getElementById('discountsM').value='';
			</script>
			</td>
		</tr>
		<tr>
			<td align="right" colspan="4">Total Discounts: -<%=scCurSign&money(discountTotalPrice+CatDiscounts)%></td>
		</tr>
	</table>
	<br />
	<table class="pcCPcontent">
		<tr>
			<th colspan="4">Other Settings&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=314')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
		</tr>
		<tr>
			<td width="18%"><strong>Shipping<input name="weightTotal" type="hidden" value="<%=pweightTotal%>"></strong></td>
			<td width="39%"><a href="javascript:;" onClick="newWindow('checkRealtimeRates.asp?idOrder=<%=pIdOrder%>&subtotal=<%=(ptotal+discountPrice-shipPrice-ptaxAmount-shipHandling-paymentPrice)%>&weight=<%=pweightTotal%>&cartQTY=<%=pTotalCartQty%>&residential=1','window2')">Check
			Real-Time rates</a></div></td>

			<% if session("adminTaxonCharges")="" then
				session("adminTaxonCharges")=pTaxonCharges
				session("adminTaxonFees")=pTaxonFees
			end if %>
			<td width="28%">
				<input name="Shipping" type="text" id="Shipping" value="<%=money(replacecommaFromDB(replace(shipPrice," ","")))%>" size="10">
				&nbsp;&nbsp;&nbsp;
				<% if pcv_IsEUMemberState=2 then %>
					<% if session("adminTaxonCharges")=1 then %> <input name="taxShipping" type="checkbox" id="taxShipping" value="YES" checked class="clearBorder">
					<% else %> <input name="taxShipping" type="checkbox" id="taxShipping" value="YES" class="clearBorder">
					<% end if %>
					&nbsp;Tax
				<% end if %>
			</td>
			<td width="15%"></td>
		</tr>
		<tr>
			<td><strong>Shipping Provider</strong></td>
			<td colspan="2">
				<input type="text" name="shippingProvider" value="<%=shipProvider%>" size="15">
				&nbsp; <strong>Shipping Method&nbsp;</strong> <input type="text" name="shippingService" value="<%=shipService%>" size="25">
				<input type="hidden" name="shippingServiceCode" value="<%=shipServiceCode%>">
			</td>
			<td><div align="right"><%=scCurSign&money(replacecommaFromDB(shipPrice))%></div></td>
		</tr>
		<tr>
			<td bgcolor="e1e1e1"><strong>Handling Charges</strong></td>
			<td bgcolor="e1e1e1">&nbsp;</td>
			<td bgcolor="e1e1e1">
				<input name="handling" type="text" id="handling" value="<%=money(replacecommaFromDB(shipHandling))%>" size="10">
				&nbsp; &nbsp;&nbsp;
				<% if pcv_IsEUMemberState=2 then %>
					<% if session("adminTaxonFees")=1 then %> <input name="taxHandling" type="checkbox" id="taxHandling" value="YES" checked class="clearBorder">
					<% else %>
					<input name="taxHandling" type="checkbox" id="taxHandling" value="YES" class="clearBorder">
					<% end if %>
					&nbsp;Tax
				<% end if %>
			</td>
			<td bgcolor="e1e1e1"><div align="right"><%=scCurSign&money(replacecommaFromDB(shipHandling))%></div></td>
		</tr>
		<tr>
			<td><strong>Payment Charges</strong></td>
			<td bgcolor="#FFFFFF">
				Calculate&nbsp;
				<input name="calculatePayPercentage" type="text" id="calculateDisPercentage" value="<%=session("adminPayPercentage")%>" size="6"> %
			</td>
			<td>
				<input name="paymentCharges" type="text" id="paymentCharges" value="<%=money(replacecommaFromDB(paymentPrice))%>" size="10">
				<input type="hidden" name="paymentType" value="<%=paymentType%>">
			</td>
			<td><div align="right"><%=scCurSign&money(replacecommaFromDB(paymentPrice))%></div></td>
		</tr>
		<%
		'// Show VAT
		if pord_VAT>0 OR pcv_IsEUMemberState<>2 then %>
			<% if pcv_IsEUMemberState=1 then %>
				<tr bgcolor="#e1e1e1">
					<td><strong>VAT</strong></td>
					<td>&nbsp;</td>
					<td>
						<input name="VATTotal" type="text" id="VATTotal" value="<%=money(pord_VAT)%>" size="10"/>
					</td>
					<td nowrap="nowrap"><div align="right">Includes <%=scCurSign&money(pord_VAT)%> of VAT</div></td>
				</tr>
			<% else %>
				<tr bgcolor="#e1e1e1">
					<td><strong>VAT</strong></td>
					<td>&nbsp;</td>
					<td>
						<input name="VATTotal" type="text" id="VATTotal" value="<%=money(0)%>" size="10"/>
					</td>
					<td nowrap="nowrap"><div align="right"><%=scCurSign&money(pord_VAT)%> of VAT Removed</div></td>
				</tr>
			<% end if %>
		<% else
			'// Show Tax
			if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
				psplitTaxType=0 %>
				<tr bgcolor="#e1e1e1">
					<td><strong>Taxes</strong></td>
					<td>Calculate&nbsp; <input name="calculateTaxPercentage" type="text" id="calculateDisPercentage" value="<%=session("adminTaxPercentage")%>" size="6">
						&nbsp; %&nbsp;&nbsp; <a href="javascript:;" onClick="newWindow('checkTaxRates.asp?S=<%=pshippingStateCode%>&C=<%=pshippingCountryCode%>&Z=<%=pshippingZip%>','window2')">Check
						Tax Rate</a></td>
					<td><input name="TaxTotal" type="text" id="TaxTotal" value="<%=money(ptaxAmount)%>" size="10"></td>
					<td><div align="right"><%=scCurSign&money(ptaxAmount)%></div></td>
				</tr>
			<% else
				psplitTaxType=1
				pTaxTypeCnt=0
				taxArray=split(ptaxDetails,",")
				for i=0 to (ubound(taxArray)-1)
					pTaxTypeCnt=pTaxTypecnt+1
					taxDesc=split(taxArray(i),"|")
					'State Taxes|1.27875,Country Taxes|0.34875,
					%>
					<tr bgcolor="#e1e1e1">
						<td><strong><%=taxDesc(0)%></strong></td>
						<td>Calculate&nbsp; <input name="calculateTaxPercentage<%=pTaxTypeCnt%>" type="text" value="<%=session("adminTaxPercentage"&pTaxTypeCnt)%>" size="6">

						&nbsp;%&nbsp;&nbsp; <a href="javascript:;" onClick="newWindow('checkTaxRates.asp?S=<%=pshippingStateCode%>&C=<%=pshippingCountryCode%>&Z=<%=pshippingZip%>&T=2','window2')">Check
						Tax Rate</a>
						</td>
						<td>
							<input type="hidden" name="TaxDesc<%=pTaxTypeCnt%>" value="<%=taxDesc(0)%>">
							<input name="TaxTotal<%=pTaxTypeCnt%>" type="text" value="<%=money(taxDesc(1))%>" size="10">
						</td>
						<td><div align="right"><%=scCurSign&money(taxDesc(1))%></div></td>
					</tr>
				<% next %>
				<input type="hidden" name="taxTypeCnt" value="<%=pTaxTypeCnt%>">
				<input type="hidden" name="splitTaxType" value="1">
			<% end if %>
		<% end if %>

		<%'GGG Add-on start
		IF pGiftCode<>"" THEN%>
			<tr>
				<td>&nbsp;</td>
				<td colspan="2"><div align="right"><strong>Gift Certificate Amount:</strong><br>(<%=pGiftCode%>)</div></td>
				<td valign="top"><div align="right"><%="-"&scCurSign&money(pGiftUsed)%></div>
				</td>
			</tr>
		<% END IF
		'GGG Add-on end%>
		<tr <%if pGiftCode<>"" then%>bgcolor="#e1e1e1"<%else%>bgcolor="#FFFFFF"<%end if%>>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td><div align="right"><strong>Total: </strong></div></td>
			<td><div align="right"><%=scCurSign&money(ptotal)%></div></td>
		</tr>
	</table>

	<table class="pcCPcontent">
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<th>Payment Status</th>
		</tr>
		<tr>
			<td>To update the payment status of this order, select status from the drop-down box below.</td>
		</tr>
		<tr>
			<td>Status:&nbsp;
				<select name="pcv_PaymentStatus">
					<option value="0" <%if pcv_PaymentStatus="0" then%>selected<%end if%>>Pending</option>
					<option value="1" <%if pcv_PaymentStatus="1" then%>selected<%end if%>>Authorized</option>
					<option value="2" <%if pcv_PaymentStatus="2" then%>selected<%end if%>>Paid</option>
					<option value="6" <%if pcv_PaymentStatus="6" then%>selected<%end if%>>Refunded</option>
					<option value="8" <%if pcv_PaymentStatus="8" then%>selected<%end if%>>Voided</option>
				</select>
			</td>
		</tr>
	</table>

<% end if %>
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="center">
		<input name="SubmitUPD" type="submit" class="submit2" id="SubmitUPD" value="Update Order">
			&nbsp;&nbsp;
		<input type="button" name="Button" value="Back to Order Details" onClick="location.href='Orddetails.asp?id=<%=qryID%>'">
		</td>
	</tr>
</table>
</form>
<%
for i=1 to pCnt
	session("adminCurrentPrice"&i)=""
	session("adminpQTY"&i)=""
next
session("adminTaxPercentage")=""
session("adminDisPercentage")=""
session("adminPayPercentage")=""

function replacecommaToDB(pricenumber)
	if (scDecSign=",") and (cdbl("3.00")<>3) then
		replacecommaToDB=replace(pricenumber,".","")
		replacecommaToDB=replace(replacecommaToDB,",",".")
	else
		if (cdbl("3.00")<>3) then
			replacecommaToDB=replace(pricenumber,",",".")
		else
			replacecommaToDB=replace(pricenumber,",","")
		end if
	end if
end function

function replacecommaFromDB(pricenumber)
	if (scDecSign=",") and (cdbl("3.00")<>3) then
		replacecommaFromDB=replace(pricenumber,".",",")
	else
		replacecommaFromDB=replace(pricenumber,",","")
	end if
end function

function replacecommaToCal(pricenumber)
	Dim tmp1
	tmp1=""
	if pricenumber<>"" then
		if cdbl("3.00")<>3 then
			if scDecSign="," then
				tmp1=replace(pricenumber,".","")
			else
				tmp1=replace(pricenumber,",","")
				tmp1=replace(tmp1,".",",")
			end if
		else
			if scDecSign="," then
				tmp1=replace(pricenumber,".","")
				tmp1=replace(tmp1,",",".")
			else
				tmp1=replace(pricenumber,",","")
			end if
		end if
		replacecommaToCal=cdbl(tmp1)
	else
		replacecommaToCal=pricenumber
	end if
end function
%>
<!--#include file="adminFooter.asp"-->