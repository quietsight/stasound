<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "addsavedprdstocart.asp"
' This page adds the saved wishlist products into the cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="chkPrices.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<%
Response.Buffer = True

Dim ItemsDiscounts, Charges, DefaultPrice, query, conntemp, rstemp, pIdOrder

'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Page On-Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
err.number=0
on error resume next

' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function

pIdOrder=getUserInput(request("idOrder"),0)
pIdOrder1=pIdOrder

ItemsDiscounts=0
Charges=0
pDefaultPrice=0

dim pcCartArray1(100,45)
if (Session("pcCartIndex")<>"") and (Session("pcCartIndex")<>"0") then
pcCartArray=session("pcCartSession")
ppcCartIndex = Session("pcCartIndex")
else

session("pcCartSession")=pcCartArray1
ppcCartIndex = 0
Session("pcCartIndex") = ppcCartIndex
pcCartArray=session("pcCartSession")
end if

call openDb()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Page On-Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td>

<%
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' START: CHECK STOCK LEVELS
'/////////////////////////////////////////////////////////////////////////////////////////////////////
IF request("OrderRepeat")<>"haveto" then

	pcv_OrdHaveOutStock=0
	pcv_OrdHaveStock=0

	'*************************************************************************************************
	' START: Build a query based off the business logic
	'*************************************************************************************************
	query="SELECT WishList.idProduct "

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: BTO
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'BTO ADDON-S
	tmpSQL1=""
	tmpSQL2=""
	If scBTO=1 then
		query=query&",idConfigWishListSession"
		if (iBTOQuoteSubmitOnly=0) then
			tmpSQL1=",products"
			tmpSQL2=" AND products.idproduct=WishList.idproduct AND products.noprices=0 AND WishList.qsubmit<1 "
		else
			tmpSQL1=""
			tmpSQL2=" AND idConfigWishListSession=0"
		end if
	End If
	'BTO ADDON-E
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: BTO
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	query=query&" FROM WishList" & tmpSQL1 & " WHERE idCustomer=" &Session("idcustomer") & tmpSQL2 & ";"
	'*************************************************************************************************
	' END: Build a query based off the business logic
	'*************************************************************************************************

	'--> execute our query
	set rstemp=conntemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

	'-->  check results for data
	if rstemp.eof then
		call closeDb()
		response.end
	end if

	'*************************************************************************************************
	' START: Loop through all of our Saved Product
	'*************************************************************************************************

	'//////////////////////////////////
	'//  Start Loop "rstemp"
	'//////////////////////////////////
	do while not rstemp.eof

		'--> Get the id of this product
		pidProduct=rstemp("idProduct")

		'--> BTO ADDON-S
		If scBTO=1 then
			pidConfig=rstemp("idConfigWishListSession")
		End if
		'--> BTO ADDON-E


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  Start: New Query with "rstemp". Do for each product!
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="SELECT DISTINCT Products.serviceSpec,Products.stock,Products.noStock,Products.pcprod_minimumqty,Products.pcprod_qtyvalidate,pcProd_multiQty,Products.pcProd_BackOrder "

		tmp1=""
		tmp2=""
		tmp3=""

		'--> BTO ADDON-S
		If scBTO=1 then
		If pidConfig<>"0" then
			tmp1=" ,ConfigWishlistSessions.pcConf_Quantity "
			tmp2=" ,ConfigWishlistSessions "
			tmp3=" AND ConfigWishlistSessions.idproduct=Products.idproduct AND ConfigWishlistSessions.idConfigWishListSession=" & pidConfig
		End If
		End if
		'--> BTO ADDON-E

		tmp2=tmp2 & ",categories,categories_products"
		tmp3=tmp3 & " AND categories_products.idproduct=products.idproduct AND categories.idcategory=categories_products.idcategory AND categories.iBTOhide=0 "
		
		if session("idCustomer")<>0 AND session("customerType")=1 then
		else
			tmp3=tmp3 & " AND categories.pccats_RetailHide=0"
		end if

		'// START v4.1 - Not For Sale override
			if NotForSaleOverride(session("customerCategory"))=1 then
				queryNFSO=""
			else
				queryNFSO=" AND products.formQuantity=0"
			end if
		'// END v4.1

		query = query & tmp1 & " from Products" & tmp2 & " WHERE products.idproduct=" & pidProduct & " AND products.removed=0 AND products.active<>0" & queryNFSO & tmp3
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  END: New Query with "rstemp". Do for each product!
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'--> execute the query
		set rs=connTemp.execute(query)

		'--> Check the data and retreive the data
		IF not rs.eof THEN


			'// set our variables if there are results
			pserviceSpec=rs("serviceSpec")
			pStock=rs("stock")
			pNoStock=rs("noStock")
			pcv_minqty=rs("pcprod_minimumqty")
			pcv_qtyvalid=rs("pcprod_qtyvalidate")
			pcv_multiQty=rs("pcProd_multiQty")
			pcv_BackOrder=rs("pcProd_BackOrder")
			if IsNull(pcv_BackOrder) or pcv_BackOrder="" then
				pcv_BackOrder=0
			end if

			if tmp1<>"" then
				pquantity=clng(rs("pcConf_Quantity"))
				if (pquantity<>"") and (pquantity>0) then
				else
				pquantity=1
				end if
			else
				pquantity=1
			end if


			if (PStock<pcv_minqty) and (pcv_qtyvalid=0) then
				pStock=0
			end if


			if pcv_qtyvalid="1" then
				if PStock<pcv_minqty then
					pStock=0
				else
					if (PStock<pquantity) and (pStock>pcv_multiQty) then
						pStock=Fix(pStock/pcv_multiQty)*pcv_multiQty
					end if
				end if
			end if


			if (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_BackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_BackOrder=0) then
				pcv_OrdHaveOutStock=1
			else
				pcv_OrdHaveStock=1
			end if

		ELSE

			'// set our variable if no results
			pcv_OrdHaveOutStock=1

		END IF

	set rs=nothing
	rstemp.MoveNext
	loop
	'//////////////////////////////////
	'//  End Loop "rstemp"
	'//////////////////////////////////

	'*************************************************************************************************
	' END: Loop through all of our Saved Product
	'*************************************************************************************************

	set rstemp=nothing


	'*************************************************************************************************
	' START: Stock Issues
	'*************************************************************************************************
	if (pcv_OrdHaveOutStock=1) and (pcv_OrdHaveStock=0) then
		call closedb()
		response.redirect "msg.asp?message=134"
	end if

	if (pcv_OrdHaveOutStock=1) and (pcv_OrdHaveStock=1) then
		call closedb()
		response.redirect "msg.asp?message=135"
	end if
	'*************************************************************************************************
	' END: Stock Issues
	'*************************************************************************************************



END IF 'request("OrderRepeat")<>"haveto"
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' END: CHECK STOCK LEVELS
'/////////////////////////////////////////////////////////////////////////////////////////////////////



'/////////////////////////////////////////////////////////////////////////////////////////////////////
' START: ADD ITEMS THAT HAVE STOCK
'/////////////////////////////////////////////////////////////////////////////////////////////////////


	'*************************************************************************************************
	' START: Build a query based off the business logic
	'*************************************************************************************************
	query="SELECT WishList.idProduct, products.btoBPrice, products.price "

	'BTO ADDON-S
	If scBTO=1 then
	query=query&",WishList.idConfigWishListSession "
		if (iBTOQuoteSubmitOnly=0) then
			tmpSQL1=""
			tmpSQL2=" AND products.noprices=0 AND WishList.qsubmit<1 "
		else
			tmpSQL1=""
			tmpSQL2=" AND WishList.idConfigWishListSession=0"
		end if
	End If
	'BTO ADDON-E

	query=query&",WishList.pcwishList_OptionsArray "

	query=query&",products.description, products.sku, products.weight, products.pcprod_QtyToPound, products.emailText, products.deliveringTime, products.pcSupplier_ID, products.cost, products.stock, products.notax, products.noshipping, products.iRewardPoints, products.pcProd_Surcharge1, products.pcProd_Surcharge2 FROM WishList, products WHERE products.idProduct=wishlist.idProduct AND WishList.idCustomer=" & Session("idcustomer") & tmpSQL2
	'*************************************************************************************************
	' END: Build a query based off the business logic
	'*************************************************************************************************

	'--> execute our query
	set rstemp=conntemp.execute(query)

	'*************************************************************************************************
	' START: Loop through all of our Saved Product
	'*************************************************************************************************

	'//////////////////////////////////
	'//  Start Loop "rstemp"
	'//////////////////////////////////
	do while not rstemp.eof

		'--> Get id of this product
		pidProduct=rstemp("idProduct")

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Product Options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcv_strSelectedOptions=""
		pcv_strSelectedOptions = rstemp("pcwishList_OptionsArray")
		pcv_strSelectedOptions=pcv_strSelectedOptions&""


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' End: Product Options
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		'--> BTO ADDON-S
		If scBTO=1 then
			pidconfig=rstemp("idConfigWishListSession")
		End if
		'--> BTO ADDON-E


		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  Start: New Query with "rstemp". Do for each product!
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query="SELECT DISTINCT products.serviceSpec,products.stock,products.noStock,products.pcprod_minimumqty,products.pcprod_qtyvalidate,products.pcProd_multiQty,Products.pcProd_BackOrder "

		tmp1=""
		tmp2=""
		tmp3=""

		'BTO ADDON-S
		If scBTO=1 then
		If pidConfig<>"0" then
			tmp1=" ,ConfigWishlistSessions.pcConf_Quantity "
			tmp2=" ,ConfigWishlistSessions "
			tmp3=" AND ConfigWishlistSessions.idproduct=Products.idproduct AND ConfigWishlistSessions.idConfigWishListSession=" & pidConfig
		End If
		End if
		'BTO ADDON-E

		tmp2=tmp2 & ",categories,categories_products"
		tmp3=tmp3 & " AND categories_products.idproduct=products.idproduct AND categories.idcategory=categories_products.idcategory AND categories.iBTOhide=0 "
		
		if session("idCustomer")<>0 AND session("customerType")=1 then
		else
			tmp3=tmp3 & " AND categories.pccats_RetailHide=0"
		end if

		'// START v4.1 - Not For Sale override
			if NotForSaleOverride(session("customerCategory"))=1 then
				queryNFSO=""
			else
				queryNFSO=" AND products.formQuantity=0"
			end if
		'// END v4.1

		query=query & tmp1 & "from Products" & tmp2 & " WHERE products.idproduct=" & pidProduct & " AND products.removed=0 AND products.active<>0" & queryNFSO & tmp3
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  End: New Query with "rstemp". Do for each product!
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		'--> execute our query
		set rs=connTemp.execute(query)


		'/////////////////////////////////////////////////////////////////////////////////////////////////
		'  Start: Add Our Items
		'/////////////////////////////////////////////////////////////////////////////////////////////////
		IF not rs.eof THEN '// Start: Add and calculate products

			'*********************************************************************************************
			'  Start: Set Additional Variables for this Product via the "rs" set
			'*********************************************************************************************
			pserviceSpec=rs("serviceSpec")
			pStock=rs("stock")
			pNoStock=rs("noStock")
			pcv_minqty=rs("pcprod_minimumqty")
			pcv_qtyvalid=rs("pcprod_qtyvalidate")
			pcv_multiQty=rs("pcProd_multiQty")
			pcv_BackOrder=rs("pcProd_BackOrder")

			if tmp1<>"" then

				pquantity=clng(rs("pcConf_Quantity"))
				if (pquantity<>"") and (pquantity>0) then
					'//
				else
					pquantity=1
				end if

			else
				pquantity=1
			end if


			if (PStock<pcv_minqty) and (pcv_qtyvalid=0) then
				pStock=0
			end if

			if pcv_qtyvalid="1" then
				if PStock<pcv_minqty then
					pStock=0
				else
					if (PStock<pquantity) and (pStock>pcv_multiQty) then
						pStock=Fix(pStock/pcv_multiQty)*pcv_multiQty
					end if
				end if
			end if
			'*********************************************************************************************
			'  End: Set Additional Variables for this Product via the "rs" set
			'*********************************************************************************************

			IF (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_BackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_BackOrder=0) THEN

				'// Do do anything, its out of stock.

			ELSE

				if pStock=0 then
					if pcv_minqty>"0" then
						PStock=pcv_minqty
					else
						pStock=1
					end if
				end if

				if pStock<pquantity then
					pquantity=pStock
				else
					if (pquantity<pcv_minqty) AND (pcv_qtyvalid=0) then
						pquantity=pcv_minqty
					else
						if (pquantity<pcv_minqty) AND (pcv_qtyvalid="1") AND (pcv_minqty>0) then
							pquantity=pcv_minqty
						else
							if (pquantity<pcv_multiQty) AND (pcv_qtyvalid="1") AND (pcv_minqty=0) then
								pquantity=pcv_multiQty
							end if
						end if
					end if
				end if

				WPrice=rstemp("btoBPrice")
				if WPrice<>"" then
				else
					WPrice="0"
				end if
				if (session("CustomerType")="1") and (WPrice<>"0") then
					punitPrice=WPrice
				else
					punitPrice=rstemp("Price")
				end if
				pDefaultPrice=punitPrice

				pxfdetails=""
				'BTO ADDON-S
				if scBTO=1 then
					pidconfigSession=pidConfig
					if pidconfigSession="0" then
						pidconfigSession=""
					end if
				End If
				'BTO ADDON-E

				pDescription = rstemp("description")
				pWeight	= rstemp("weight")
				pcv_QtyToPound=rstemp("pcprod_QtyToPound")
				if pcv_QtyToPound>0 then
					pWeight=(16/pcv_QtyToPound)
					if scShipFromWeightUnit="KGS" then
						pWeight=(1000/pcv_QtyToPound)
					end if
				end if
				pSku = rstemp("sku")
				pEmailText = rstemp("emailText")
				pDeliveringTime	= rstemp("deliveringTime")
				pIdSupplier	= rstemp("pcSupplier_ID")
				pCost = rstemp("cost")
				pStock = rstemp("stock")
				pnotax = rstemp("notax")
				pnoshipping = rstemp("noshipping")
				iRewardPoints=rstemp("iRewardPoints")
				pcv_Surcharge1 = rstemp("pcProd_Surcharge1")
				pcv_Surcharge2 = rstemp("pcProd_Surcharge2")

				pIdConfigSession=trim(pidconfigSession)


			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START:  Get the Options for the item
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))

			pcv_strOptionsArray = ""
			pcv_strOptionsPriceArray = ""
			pcv_strOptionsPriceArrayCur = ""
			pcv_strOptionsPriceTotal = 0
			xOptionsArrayCount = 0
			pPriceToAdd = 0

			For cCounter = LBound(pcArray_SelectedOptions) TO UBound(pcArray_SelectedOptions)

				' SELECT DATA SET
				' TABLES: optionsGroups, options, options_optionsGroups
				query = 		"SELECT optionsGroups.optionGroupDesc, options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice "
				query = query & "FROM optionsGroups, options, options_optionsGroups "
				query = query & "WHERE idoptoptgrp=" & pcArray_SelectedOptions(cCounter) & " "
				query = query & "AND options_optionsGroups.idOption=options.idoption "
				query = query & "AND options_optionsGroups.idOptionGroup=optionsGroups.idoptiongroup "

				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rs=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if


				if rs.eof then
					set rs=nothing
					call closeDb()
					response.redirect "msg.asp?message=42"
				else

					xOptionsArrayCount = xOptionsArrayCount + 1

					pOptionDescrip=""
					pOptionGroupDesc=""
					pPriceToAdd=""
					pOptionDescrip=rs("optiondescrip")
					pOptionGroupDesc=rs("optionGroupDesc")

					If Session("customerType")=1 Then
						pPriceToAdd=rs("Wprice")
						If rs("Wprice")=0 then
							pPriceToAdd=rs("price")
						End If
					Else
						pPriceToAdd=rs("price")
					End If

					'// Generate Our Strings
					if xOptionsArrayCount > 1 then
						pcv_strOptionsArray = pcv_strOptionsArray & chr(124)
						pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & chr(124)
						pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & chr(124)
					end if
					'// Column 4) This is the Array of Product "option groups: options"
					pcv_strOptionsArray = pcv_strOptionsArray & pOptionGroupDesc & ": " & pOptionDescrip
					'// Column 25) This is the Array of Individual Options Prices
					pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & pPriceToAdd
					'// Column 26) This is the Array of Individual Options Prices, but stored as currency "scCurSign & money(pcv_strOptionsPriceTotal) "
					pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & money(pPriceToAdd)
					'// Column 5) This is the total of all option prices
					pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd

				end if

				set rs=nothing
			Next
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END:  Get the Options for the item
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Create new Product Config Session
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

				if pIdConfigSession<>"" then
					query="select * from configWishListSessions where IdConfigWishListSession=" & pIdConfigSession
					set rs=conntemp.execute(query)

					IF not rs.eof THEN
						PreRecord=""
						PreRecord1=""
						pConfigKey=trim(randomNumber(9999)&randomNumber(9999))

						iCols = rs.Fields.Count
						for dd=1 to iCols-1
						IF (Ucase(Rs.Fields.Item(dd).Name)="IDOPTIONA") OR (Ucase(Rs.Fields.Item(dd).Name)="IDOPTIONB") OR (Ucase(Rs.Fields.Item(dd).Name)="XFDETAILS") OR (Ucase(Rs.Fields.Item(dd).Name)="FPRICE") OR (Ucase(Rs.Fields.Item(dd).Name)="DPRICE") OR (Ucase(Rs.Fields.Item(dd).Name)="PCCONF_QUANTITY") OR (Ucase(Rs.Fields.Item(dd).Name)="PCCONF_QDISCOUNT") OR (Ucase(Rs.Fields.Item(dd).Name)="QDISCOUNTS") THEN
						ELSE
						if dd=1 then
							PreRecord=PreRecord & Rs.Fields.Item(dd).Name
						else
							PreRecord=PreRecord & "," & Rs.Fields.Item(dd).Name
						end if
						IF Ucase(Rs.Fields.Item(dd).Name)="CONFIGKEY" then
								if dd=1 then
									PreRecord1=PreRecord1 & "'" & pConfigKey & "'"
								else
									PreRecord1=PreRecord1 & ",'" & pConfigKey & "'"
								end if
							ELSE
								IF Ucase(Rs.Fields.Item(dd).Name)="DTCREATED" then
									if scDB="Access" then
										myStr11="#"
									else
										myStr11="'"
									end if
									if dd=1 then
										PreRecord1=PreRecord1 & myStr11 & Now() & myStr11
									else
										PreRecord1=PreRecord1 & "," & myStr11 & Now() & myStr11
									end if
								ELSE
									IF Ucase(Rs.Fields.Item(dd).Name)="STRINGOPTIONS" then
										if dd=1 then
											PreRecord1=PreRecord1 & "' '"
										else
											PreRecord1=PreRecord1 & ",' '"
										end if
									ELSE
										FType="" & Rs.Fields.Item(dd).Type
										if (Ftype="202") or (Ftype="203") then
											PTemp=Rs.Fields.Item(dd).Value
											if PTemp<>"" then
												PTemp=replace(PTemp,"'","''")
											end if
											if dd=1 then
												PreRecord1=PreRecord1 & "'" & PTemp & "'"
											else
												PreRecord1=PreRecord1 & ",'" & PTemp & "'"
											end if
										else
											PTemp="" & Rs.Fields.Item(dd).Value
											if PTemp<>"" then
											else
												PTemp="0"
											end if
											if dd=1 then
												PreRecord1=PreRecord1 & PTemp
											else
												PreRecord1=PreRecord1 & "," & PTemp
											end if
										end if
									END IF 'stringOptions
								END IF 'DTCreated
							END IF 'Config Key
							END IF 'Not available fields
						next

						query="insert into configSessions (" & PreRecord & ") values (" & PreRecord1 & ")"
						set rs=conntemp.execute(query)

						query="select idConfigSession from configSessions order by idConfigSession desc"
						set rs=conntemp.execute(query)
						pIdConfigSession=rs("IdConfigSession")
						if pIdConfigSession<>"0" then
							punitPrice=updPrices(pidProduct,pIdConfigSession)
						end if
					END IF
				end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: Create new Product Config Session
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' START: Calculate BTO Items Weights
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			IF pIdConfigSession<>"" then
				query="SELECT stringProducts,stringCProducts,stringQuantity FROM configSessions WHERE idconfigSession=" & pIdConfigSession
				set rs=conntemp.execute(query)

				stringProducts=rs("stringProducts")
				stringCProducts=rs("stringCProducts")
				ArrProduct=Split(stringProducts, ",")
				ArrCProduct=Split(stringCProducts, ",")
				Qstring=rs("stringQuantity")
				ArrQuantity=Split(Qstring,",")

				CWeight=0

				if ArrProduct(0)<>"na" then
					for j=lbound(ArrProduct) to (UBound(ArrProduct)-1)
						query="SELECT weight,pcprod_QtyToPound FROM products WHERE IDProduct=" & ArrProduct(j)
						set rs=conntemp.execute(query)
						if not rs.eof then
							ItemWeight=rs("weight")
							pcv_QtyToPound=rs("pcprod_QtyToPound")
							if pcv_QtyToPound>0 then
								ItemWeight=cdbl(16/pcv_QtyToPound)
								if scShipFromWeightUnit="KGS" then
									ItemWeight=cdbl(1000/pcv_QtyToPound)
								end if
							end if

							CWeight=CWeight+cdbl(ItemWeight*clng(ArrQuantity(j)))
						end if
					next
				end if

				if ArrCProduct(0)<>"na" then
					for j=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
						query="SELECT weight,pcprod_QtyToPound FROM products WHERE IDProduct=" & ArrCProduct(j)
						set rs=conntemp.execute(query)

						if not rs.eof then

							ItemWeight=rs("weight")

							pcv_QtyToPound=rs("pcprod_QtyToPound")
							if pcv_QtyToPound>0 then
								ItemWeight=(16/pcv_QtyToPound)
								if scShipFromWeightUnit="KGS" then
									ItemWeight=(1000/pcv_QtyToPound)
								end if
							end if

							CWeight=CWeight+cdbl(ItemWeight)

						end if
					next
				end if

			END IF 'Have ID Config Session
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' END: Calculate BTO Items Weights
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			lineNumber=0
			dim iNextIndex, iCapturedNext
			iNextIndex=cint(0)
			iCapturedNext=cint(0)

			'/////////////////////////////////////////
			'// START CART LOOP
			'/////////////////////////////////////////
			for f=1 to ppcCartIndex

					'check if index is deleted and then flag
					if pcCartArray(f,10)=1 AND iCapturedNext=0 then
						iNextIndex=f
						iCapturedNext=1
					end if
					' Start: if item is not deleted and the idProduct=idProduct of added item
					if (pcCartArray(f,10)=0) and (Int(pcCartArray(f,0))=Int(trim(pIdProduct))) and (pcCartArray(f,16) & ""=trim(pIdConfigSession)) then
						if scOutofstockpurchase=-1 AND pNoStock=0 AND pcv_BackOrder=0 then
							iTempStockTotal=0
							for g=1 to ppcCartIndex
								if (pcCartArray(g,10)=0) and (pcCartArray(g,0)=trim(pIdProduct)) then
									'checking stock level
									iTempStockTotal=Int(pcCartArray(g,2))+Int(iTempStockTotal)
								end if
							next
							iTempStockTotal=Int(iTempStockTotal)+Int(pQuantity)
							if Int(iTempStockTotal)>Int(pStock) then

								'Set session variables to handle error on msg.asp
								session("pcErrStrPrdDesc") = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)
								session("pcErrIntStock") = pStock
								response.redirect "msg.asp?message=204"
							end if
						end if

						if xfieldsCnt=0 then

							'********************************************
							'// UPDATE CONDITIONS
							'********************************************

							'// check item with no optionals
							if trim(pcv_strOptionsArray)="" AND trim(pcCartArray(f,4))="" then
								 lineNumber=f
							end if

							' check only optionals with values (previously check A and B)
							if trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions) AND trim(pcv_strSelectedOptions)<>"" then
								 lineNumber=f
							end if

							' check any optionals
							if trim(pcCartArray(f,11))=trim(pcv_strSelectedOptions) then
								 lineNumber=f
							end if

						end if '// if xfieldsCnt=0 then

					end if
					' End: if item is not deleted and the idProduct=idProduct of added item
			next
			'/////////////////////////////////////////
			'// END CART LOOP
			'/////////////////////////////////////////

			Dim checkSS
			if lineNumber=0 then
				if scOutofstockpurchase=-1 AND pNoStock=0 AND pcv_BackOrder=0 then
					pTotalQuantity=pQuantity
					If Int(pQuantity)<Int(pcv_minqty) then
						pQuantity=pcv_minqty
					End if
					if Int(pQuantity)>Int(pStock) then
						pQuantity=Int(pStock)
					end if
				end if
				' insert new basket line, is not in the cart
				if pQuantity <=Int(scAddLimit) then
					pTotalQuantity=pQuantity
				else
					pQuantity=Int(scAddLimit)
					pTotalQuantity=pQuantity
				end if

				if iCapturedNext=1 then
					ppcCartIndex=iNextIndex
				else
					if tIndex<>0 then
						ppcCartIndex=tIndex
					else
						ppcCartIndex=ppcCartIndex + 1
						err.clear
						checkSS=pcCartArray(ppcCartIndex, 0)
						if err.number<>0 then
							if instr(ucase(err.description, "")) then
								response.write dictLanguage.Item(Session("language")&"_instPrd_1")
								response.End()
							end if
						end if
						session("pcCartIndex")=ppcCartIndex
					end if
				end if


				session("pcCartIndex")	= ppcCartIndex

				pcCartArray(ppcCartIndex,0) = pIdProduct
				pcCartArray(ppcCartIndex,1) = pDescription
				pcCartArray(ppcCartIndex,2) = pQuantity

				' add price or BtoB price depending on customer type
				pcCartArray(ppcCartIndex,3) = punitPrice

				pcCartArray(ppcCartIndex,8)="" '// not in use anymore


				if len(pcv_strOptionsArray)>0 then
					 pcCartArray(ppcCartIndex,4)= pcv_strOptionsArray '// store array of product "option groups: options"
				else
					pcCartArray(ppcCartIndex,4)=""
				end if

				pcCartArray(ppcCartIndex,5)= pcv_strOptionsPriceTotal '// Total Cost of all Options
				pcCartArray(ppcCartIndex,23)=pOverSizeSpec
				pcCartArray(ppcCartIndex,25)=pcv_strOptionsPriceArray '// Array of Individual Options Prices
				pcCartArray(ppcCartIndex,26)= pcv_strOptionsPriceArrayCur '// Array of Options Prices stored as currency  'scCurSign & money(pPriceToAdd)
				pcCartArray(ppcCartIndex,27)="" '// Not in use anymore
				pcCartArray(ppcCartIndex,28)="" '// Not in use anymore
				pcCartArray(ppcCartIndex,6) = pWeight + Cweight
				pcCartArray(ppcCartIndex,7) = pSku
				pcCartArray(ppcCartIndex,9) = pDeliveringTime

				' deleted mark
				pcCartArray(ppcCartIndex,10) = 0

				pcCartArray(ppcCartIndex,11)=pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers
				pcCartArray(ppcCartIndex,12)="" '// Not in use anymore
				pcCartArray(ppcCartIndex,13) = pIdSupplier
				pcCartArray(ppcCartIndex,14) = pCost
				pcCartArray(ppcCartIndex,16) = pIdConfigSession
				pcCartArray(ppcCartIndex,19) = pnotax
				pcCartArray(ppcCartIndex,20) = pnoshipping
				pcCartArray(ppcCartIndex,21) = pxfdetails
				pcCartArray(ppcCartIndex,36) = pcv_Surcharge1
				pcCartArray(ppcCartIndex,37) = pcv_Surcharge2

				pTotalQuantity = pQuantity

				'-----------------------------
				'ReCalculate BTO Items Discounts

				pcCartArray(ppcCartIndex,30)=cdbl(ItemsDiscounts)

				'End ReCulculate BTO Items Discounts
				'------------------------------------

				'------------------------------------
				'BTO Additional Charges

				pcCartArray(ppcCartIndex,31) = Cdbl(Charges)

				'End BTO Additional Charges
				'------------------------------------

			else 'Existing Products in ProductCart
				' item is already in the cart
				If pcCartArray(lineNumber,2)+ Int(pQuantity)<Int(pcv_minqty) then
						pcCartArray(lineNumber,2)=Int(pcv_minqty)
						pTotalQuantity=pcCartArray(lineNumber,2)
				else
				if pcCartArray(lineNumber,2)+ Int(pQuantity) <=Int(scAddLimit) then
					' quantity added + previous quantity is not more than allowed
					pcCartArray(lineNumber,2)=Int(pcCartArray(lineNumber,2)) + Int(pQuantity)
					pTotalQuantity=pcCartArray(lineNumber,2)
				else
					pcCartArray(lineNumber,2)=Int(scAddLimit)
					pTotalQuantity=pcCartArray(lineNumber,2)
				end if
				end if
			end if

				' get discount per quantity
				query="SELECT * FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pTotalQuantity& " AND quantityUntil>=" &pTotalQuantity
				set rstemp1=conntemp.execute(query)

				if err.number <> 0 and err.number<>9 then
					call LogErrorToDatabase()
					set rstemp1=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				if Session("customerType")=1 OR session("customerCategory")<>0 then
					pcCartArray(f,18)=1
				else
					pcCartArray(f,18)=0
				end if

				pOrigPrice = pcCartArray(ppcCartIndex,3)
				pcCartArray(ppcCartIndex,17) = pOrigPrice
				if pcQDiscountType<>"1" then
				pOrigPrice=pOrigPrice+pcCartArray(ppcCartIndex,5)+(pcCartArray(ppcCartIndex,30)/pTotalQuantity)
				else
				pOrigPrice=pDefaultPrice
				end if

				pcCartArray(ppcCartIndex,15) = 0
				if not rstemp1.eof and err.number<>9 then
					' there are quantity discounts defined for that quantity
					pDiscountPerUnit = rstemp1("discountPerUnit")
					pDiscountPerWUnit = rstemp1("discountPerWUnit")
					pPercentage = rstemp1("percentage")

					if session("customerType")<>1 then
						if pPercentage = "0" then
							pcCartArray(ppcCartIndex,3)  = pcCartArray(ppcCartIndex,3) - pDiscountPerUnit
							pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (pDiscountPerUnit * pTotalQuantity)
						else
							pcCartArray(ppcCartIndex,3) = pcCartArray(ppcCartIndex,3) - ((pDiscountPerUnit/100) * pOrigPrice)
							pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + ((pDiscountPerUnit/100) * (pOrigPrice * pTotalQuantity))
						end if
					else
						if pPercentage = "0" then
							pcCartArray(ppcCartIndex,3)  = pcCartArray(ppcCartIndex,3) - pDiscountPerWUnit
							pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + (pDiscountPerWUnit * pTotalQuantity)
						else
							pcCartArray(ppcCartIndex,3) = pcCartArray(ppcCartIndex,3) - ((pDiscountPerWUnit/100) * pOrigPrice)
							pcCartArray(ppcCartIndex,15) = pcCartArray(ppcCartIndex,15) + ((pDiscountPerWUnit/100) * (pOrigPrice * pTotalQuantity))
						end if
					end if
				end if

					'// Start Reward Points
					If RewardsActive <> 0 Then
						pcCartArray(ppcCartIndex,22) = Clng(iRewardPoints)
					End If
					'// End Reward Points

			END IF '// IF (scOutofStockPurchase=-1...

		END IF '// End: Add and calculate products
		'/////////////////////////////////////////////////////////////////////////////////////////////////
		'  End: Add Our Items
		'/////////////////////////////////////////////////////////////////////////////////////////////////


	rstemp.movenext
	loop
	'//////////////////////////////////
	'//  End Loop "rstemp"
	'//////////////////////////////////

	'*************************************************************************************************
	' END: Loop through all of our Saved Product
	'*************************************************************************************************

'/////////////////////////////////////////////////////////////////////////////////////////////////////
' END: ADD ITEMS THAT HAVE STOCK
'/////////////////////////////////////////////////////////////////////////////////////////////////////

%>
<!--#include file="inc-UpdPrdQtyDiscounts.asp"-->
<%
pcCartArray(1,18)=0
%>
<!--#include file="pcReCalPricesLogin.asp"-->
<%
session("pcCartSession") = pcCartArray
Session("pcCartIndex") = ppcCartIndex

call closedb()

response.redirect "viewcart.asp"
%>

		Please wait while we process your items...

		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->