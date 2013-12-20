<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "cRec.asp"
' This page changes quantities of items in cart and recalculate the cart values and totals.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="inc_checkPrdQtyCart.asp"-->
<!--#include file="inc_checkMinMul.asp"-->
<%
Response.Buffer = True

Dim query, conntemp, rstemp, pcCartArray, f, pNewQuantity, pIdProduct, tempQty

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

'arrService=Session("Service")
'pServiceIndex=Session("serviceIndex")

if int(countCartRows(pcCartArray, pcCartIndex))=0 then
	response.redirect "msg.asp?message=16" 
end if

'GGG Add-on start
if request("actGW")<>"" then
	for f=1 to pcCartIndex
		GW=getUserInput(request("GW" & f),0)
		if GW<>"" then
			if (pcCartArray(f,34)<>"") and (pcCartArray(f,34)<>"0") then
			else
				pcCartArray(f,34)=GW
			end if
		else
			pcCartArray(f,34)=""
		end if
	Next

	session("pcCartSession")=pcCartArray
	dim strGwRedirectSSL
	strGwRedirectSSL="onepagecheckout.asp"
	if scSSL="1" AND scIntSSLPage="1" then
		strGwRedirectSSL=replace((scSslURL&"/"&scPcFolder&"/pc/onepagecheckout.asp"),"//","/")
		strGwRedirectSSL=replace(strGwRedirectSSL,"https:/","https://")
		strGwRedirectSSL=replace(strGwRedirectSSL,"http:/","http://")
	end if
	response.redirect strGwRedirectSSL
else
	if request("actGCs")<>"" then
		dim strNrRedirectSSL
		strNrRedirectSSL="onepagecheckout.asp"
		if scSSL="1" AND scIntSSLPage="1" then
			strNrRedirectSSL=replace((scSslURL&"/"&scPcFolder&"/pc/onepagecheckout.asp"),"//","/")
			strNrRedirectSSL=replace(strNrRedirectSSL,"https:/","https://")
			strNrRedirectSSL=replace(strNrRedirectSSL,"http:/","http://")
		end if
	response.redirect strNrRedirectSSL
	end if
end if
'GGG Add-on end

' insert new quantity
for f=1 to pcCartIndex

	if pcCartArray(f,10)=0 then
		if Instr(session("sf_FQuotes"),"****" & pcCartArray(f,0) & "****")=0 then 'Does not update Finalized Quotes
			'GGG Add-on start
			GW=getUserInput(request("GW" & f),0)
			if GW<>"" then
				if (pcCartArray(f,34)<>"") and (pcCartArray(f,34)<>"0") then
				else
					pcCartArray(f,34)=GW
				end if
			else
				pcCartArray(f,34)=""
			end if
			'GGG Add-on end
			' identity which index item 
			tempQty=trim(request.Form("Cant" & Cstr(f)))
			if NOT validNum(tempQty) then
				tempQty=1
			end if
			  
			'//  Bundled Child
			if (pcCartArray(f,27)>"0") AND (pcCartArray(f,12)="0") then  
					'// Changed below...
					tempQty = pcCartArray(cint(pcCartArray(f,27)),2)
			end if
			
			'// Required Accessory to Parent
			if (pcCartArray(f,27)>"0") AND (pcCartArray(f,12)="-2") then
					'// Changed below...
					tempQty = pcCartArray(cint(pcCartArray(f,27)),2)
			end if
			
		if int(tempQty)<>int(pcCartArray(f,2)) then 
			pIdProduct=pcCartArray(f,0)
			pNewQuantity=tempQty
			if NOT validNum(pNewQuantity) OR int(pNewQuantity)<1 then
				pNewQuantity=1
			end if
			if scOutofstockpurchase=-1 then
				call opendb()
				'Check Product Stock
				queryC="SELECT idProduct,stock,description FROM Products WHERE idProduct=" & pIdProduct & " AND (nostock=0) AND (pcProd_BackOrder=0);"
				set rsC=ConnTemp.execute(queryC)
				if not rsC.eof then
					tmpID=rsC("idProduct")
					tmpiStock=rsC("stock")
					tmpiDesc=rsC("description")
					set rsC=nothing
					if CheckOFS(tmpID,pNewQuantity,tmpiStock,f)=1 then
						call closedb()
						response.Clear()
						response.redirect "msgb.asp?message="&Server.Urlencode("The quantity of "&tmpiDesc&" that you are trying to order is greater than the quantity that we currently have in stock. We currently have "&tmpiStock&" unit(s) in stock.<br><br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &""" border=0></a>" )
					end if
				end if
				set rsC=nothing
				
				call CheckMinMulQty(pIdProduct,pNewQuantity)
				
				'Check BTO Items stock
				if (pcCartArray(f,16)<>"") AND (pcCartArray(f,16)>"0") then
					if scOutofstockpurchase=-1 AND iBTOOutofstockpurchase=-1 then
						queryC="SELECT stringProducts, stringQuantity FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16)) & ";"
						set rsC=connTemp.execute(queryC)
						if not rsC.eof then
							tmpCSP=split(rsC("stringProducts"),",")
							tmpCSQ=split(rsC("stringQuantity"),",")
							set rsC=nothing
							For cl=lbound(tmpCSP) to ubound(tmpCSP)
								if trim(tmpCSP(cl))<>"" then
									queryC="SELECT idProduct,stock,description FROM Products WHERE idProduct=" & tmpCSP(cl) & " AND (nostock=0) AND (pcProd_BackOrder=0);"
									set rsC=ConnTemp.execute(queryC)
									if not rsC.eof then
										tmpID=rsC("idProduct")
										tmpiStock=rsC("stock")
										tmpiDesc=rsC("description")
										set rsC=nothing
										if CheckOFS(tmpID,Clng(tmpCSQ(cl))*pNewQuantity,tmpiStock,f)=1 then
											call closedb()
											response.Clear()
											response.redirect "msgb.asp?message="&Server.Urlencode("The quantity of "&tmpiDesc&" that you are trying to order is greater than the quantity that we currently have in stock. We currently have "&tmpiStock&" unit(s) in stock.<br><br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &""" border=0></a>" )
										end if
									end if
									set rsC=nothing
									call CheckMinMulQty(Clng(tmpCSQ(cl)),Clng(tmpCSQ(cl))*pNewQuantity)
								end if
							Next
						end if
						set rsC=nothing
					end if
				end if
				call closedb()
			end if	
			pcCartArray(f,2)=pNewQuantity
			if int(pNewQuantity)>int(scAddLimit) then
				response.redirect "msg.asp?message=17"                   
			end if
			
			' check discounts per quantity and recalculates the price 
			'
			if pIdProduct<>"" and pcCartArray(f,16)="" then
			
				call opendb()

				'====================
				' get original price
				'====================
				query="SELECT description, price, bToBPrice, stock, noStock FROM products WHERE idProduct=" &pIdProduct
				set rstemp=conntemp.execute(query)
				
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
	 
				if rstemp.eof then
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Items do not exist") 					
				end if
				
				pDescription=rstemp("description")
				tempprice=rstemp("price")
				tempbToBPrice=rstemp("bToBPrice")
				pStock=rstemp("stock")
				pNoStock=rstemp("noStock")
				
				set rstemp = nothing
				
				'Check if this customer is logged in with a customer category
				if session("customerCategory")<>0 then
					query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					if NOT rs.eof then
						strcustomerCategory="YES"
						dblpcCC_Price=rs("pcCC_Price")
						dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
					else
						strcustomerCategory="NO"
					end if
					
					set rs=nothing
				end if
					
				if (tempbToBPrice<>0) then
					tempBPrice=tempbToBPrice
				else
					tempBPrice=tempprice
				end if

				if session("customerType")=1 then
					pPrice=tempBPrice
				else
					pPrice=tempprice
				end if 

				if session("customerCategoryType")="ATB" then
					if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
						tempBPrice=tempBPrice-(pcf_Round(tempBPrice*(cdbl(session("ATBPercentage"))/100),2))
						pPrice=tempBPrice
					end if
					if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
						tempprice=tempprice-(pcf_Round(tempprice*(cdbl(session("ATBPercentage"))/100),2))
						pPrice=tempprice
					end if					
				end if
				
				'if strcustomerCategory="YES" AND dblpcCC_Price>0 then
				if strcustomerCategory="YES" then
					pcCartArray(ppcCartIndex,3)=dblpcCC_Price
					pPrice=dblpcCC_Price
				end if			
				
				
				'*************************************************************************************************
				' START: GET OPTIONS
				'*************************************************************************************************
					
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START:  Get the Options for the item
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				pcArray_SelectedOptions = Split(pcCartArray(ppcCartIndex,11),chr(124))
				
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
						set rstemp=nothing
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
				
				pcCartArray(ppcCartIndex,25)=pcv_strOptionsPriceArray '// Array of Individual Options Prices
				pcCartArray(ppcCartIndex,26)= pcv_strOptionsPriceArrayCur
				'pcCartArray(ppcCartIndex,27)="" '// Not in use anymore - VERIFY FOR Crosssell
				'pcCartArray(ppcCartIndex,28)="" '// Not in use anymore - VERIFY FOR Crosssell
				pcCartArray(ppcCartIndex,5)= pcv_strOptionsPriceTotal '// Total Cost of all Options
							
				'*************************************************************************************************
				' END: GET OPTIONS
				'*************************************************************************************************
				
				
				' get discount per quantity
				query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<="&Int(pNewQuantity)&" AND quantityUntil>="&Int(pNewQuantity) 
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rstemp=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				pcCartArray(f,15)=Cint(0)
				pOrigPrice=pPrice
				pcCartArray(f,17)=pOrigPrice
				if not rstemp.eof then
					' there are quantity discounts defined for that quantity
					pDiscountPerUnit=rstemp("discountPerUnit")
					pDiscountPerWUnit=rstemp("discountPerWUnit")
					pPercentage=rstemp("percentage")
					pbaseproductonly=rstemp("baseproductonly")
	
					if session("customerType")<>1 then
						if pPercentage="0" then 
							pcCartArray(f,3)=pPrice - pDiscountPerUnit 
							pcCartArray(f,15)=pcCartArray(f,15) + (pDiscountPerUnit * pNewQuantity)
						else
							if pbaseproductonly="-1" then
								pcCartArray(f,3)=pPrice - ((pDiscountPerUnit/100) * pcCartArray(f,17))
							else
								pcCartArray(f,3)=pPrice - ((pDiscountPerUnit/100) * (pcCartArray(f,17)+pcCartArray(f,5)))
							end if
							if pbaseproductonly="-1" then
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerUnit/100) * pOrigPrice) * pNewQuantity)
							else
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerUnit/100) * (pOrigPrice+pcCartArray(f,5))) * pNewQuantity)
							end if
						end if
					else
						if pPercentage="0" then 
							pcCartArray(f,3)=pPrice - pDiscountPerWUnit
							pcCartArray(f,15)=pcCartArray(f,15) + (pDiscountPerWUnit * pNewQuantity)
						else
							if pbaseproductonly="-1" then
								pcCartArray(f,3)=pPrice - ((pDiscountPerWUnit/100) * pcCartArray(f,17))
							else
								pcCartArray(f,3)=pPrice - ((pDiscountPerWUnit/100) * (pcCartArray(f,17)+pcCartArray(f,5)))
							end if
							if pbaseproductonly="-1" then
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerWUnit/100) * pOrigPrice)* pNewQuantity)
							else
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerWUnit/100) * (pOrigPrice+pcCartArray(f,5))) * pNewQuantity)
							end if
						end if
					end if
				else
					pcCartArray(f,3)=pPrice
				end if
				set rstemp=nothing
			else 'pcCartArray(f,16)=""
			
				call opendb()
			
				'***************************************
				' START: Recalculate BTO Item Discounts
				'***************************************

				ItemsDiscounts=0
				pTotalQuantity=pNewQuantity
				
					IF pcCartArray(f,16)<>"0" and pcCartArray(f,16)<>"" then
						query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pcCartArray(f,16)
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
					
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					
						if not rstemp.eof then
							Pstring=rstemp("stringProducts")
							Vstring=rstemp("stringValues")
							Cstring=rstemp("stringCategories")
							Qstring=rstemp("stringQuantity")
							Pricestring=rstemp("stringPrice")
						end if
						
						set rstemp=nothing
						
						if Pstring<>"" and uCase(Pstring)<>"NA" then
						
							ArrProduct=split(Pstring,",")
							ArrValue=Split(Vstring, ",")
							ArrCategory=Split(Cstring, ",")
							ArrQuantity=Split(Qstring,",")
							ArrPrice=split(Pricestring,",")
							
							For m=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								'====================
								' get item discounts
								'====================
								query="SELECT quantityFrom,quantityUntil,discountperUnit,percentage,discountperWUnit FROM discountsPerQuantity WHERE IDProduct=" & ArrProduct(m)
								set rstemp=connTemp.execute(query)

								if err.number<>0 then
									call LogErrorToDatabase()
									set rstemp=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
 
								TempDiscount=0
								do while not rstemp.eof
					 				QFrom=rstemp("quantityFrom")
									QTo=rstemp("quantityUntil")
									DUnit=rstemp("discountperUnit")
									QPercent=rstemp("percentage")
									DWUnit=rstemp("discountperWUnit")
									if (DWUnit=0) and (DUnit>0) then
										DWUnit=DUnit
									end if
									

									TempD1=0
									if (clng(ArrQuantity(m)*pTotalQuantity)>=clng(QFrom)) and (clng(ArrQuantity(m)*pTotalQuantity)<=clng(QTo)) then
										if QPercent="-1" then
											if session("customerType")=1 then
												TempD1=ArrQuantity(m)*pTotalQuantity*ArrPrice(m)*0.01*DWUnit
											else
												TempD1=ArrQuantity(m)*pTotalQuantity*ArrPrice(m)*0.01*DUnit
											end if
										else
											if session("customerType")=1 then
												TempD1=ArrQuantity(m)*pTotalQuantity*DWUnit
											else
												TempD1=ArrQuantity(m)*pTotalQuantity*DUnit
											end if
										end if
									end if
									TempDiscount=TempDiscount+TempD1
									rstemp.movenext
								loop
								set rstemp=nothing
								ItemsDiscounts=ItemsDiscounts+TempDiscount
							Next

						end if 'Have BTO Items
						
					END IF
				
				pcCartArray(f,30)=ItemsDiscounts
					
				'***************************************
				' END: Recalculate BTO Item Discounts
				'***************************************
				
				IF pcQDiscountType="1" THEN
				'====================
				' get original price
				'====================
				query="SELECT price, bToBPrice FROM products WHERE idProduct=" &pIdProduct
				set rstemp=conntemp.execute(query)
				
				if not rstemp.eof then			
					tempprice=rstemp("price")
					tempbToBPrice=rstemp("bToBPrice")
				end if
				
				set rstemp = nothing
				
				'Check if this customer is logged in with a customer category
				if session("customerCategory")<>0 then
					query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					if NOT rs.eof then
						strcustomerCategory="YES"
						dblpcCC_Price=rs("pcCC_Price")
						dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
					else
						strcustomerCategory="NO"
					end if
					
					set rs=nothing
				end if
					
				if (tempbToBPrice<>0) then
					tempBPrice=tempbToBPrice
				else
					tempBPrice=tempprice
				end if

				if session("customerType")=1 then
					pPrice=tempBPrice
				else
					pPrice=tempprice
				end if 

				if session("customerCategoryType")="ATB" then
					if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
						tempBPrice=tempBPrice-(pcf_Round(tempBPrice*(cdbl(session("ATBPercentage"))/100),2))
						pPrice=tempBPrice
					end if
					if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
						tempprice=tempprice-(pcf_Round(tempprice*(cdbl(session("ATBPercentage"))/100),2))
						pPrice=tempprice
					end if					
				end if
				
				'if strcustomerCategory="YES" AND dblpcCC_Price>0 then
				if strcustomerCategory="YES" then
					pcCartArray(ppcCartIndex,3)=dblpcCC_Price
					pPrice=dblpcCC_Price
				end if
				
				END IF 'pcQDiscountType="1"
			
				' get discount per quantity
				query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<="&Int(pNewQuantity)&" AND quantityUntil>="&Int(pNewQuantity) 
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rstemp=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				pcCartArray(f,15)=Cint(0)
				pOrigPrice=pcCartArray(f,17)
				if pcQDiscountType<>"1" then
					pOrigPrice=pOrigPrice-(pcCartArray(f,30)/pNewQuantity)
				else
					pOrigPrice=pPrice
				end if
				
				if not rstemp.eof then
					' there are quantity discounts defined for that quantity
					pDiscountPerUnit=rstemp("discountPerUnit")
					pDiscountPerWUnit=rstemp("discountPerWUnit")
					pPercentage=rstemp("percentage")
					pbaseproductonly=rstemp("baseproductonly")
	
					if session("customerType")<>1 then
						if pPercentage="0" then 
							pcCartArray(f,15)=pcCartArray(f,15) + (pDiscountPerUnit * pNewQuantity)
						else
							if pbaseproductonly="-1" then
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerUnit/100) * pOrigPrice) * pNewQuantity)
							else
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerUnit/100) * (pOrigPrice+pcCartArray(f,5))) * pNewQuantity)
							end if
						end if
					else
						if pPercentage="0" then 
						pcCartArray(f,15)=pcCartArray(f,15) + (pDiscountPerWUnit * pNewQuantity)
						else
							if pbaseproductonly="-1" then
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerWUnit/100) * pOrigPrice)* pNewQuantity)
							else
								pcCartArray(f,15)=pcCartArray(f,15) + (((pDiscountPerWUnit/100) * (pOrigPrice+pcCartArray(f,5))) * pNewQuantity)
							end if
						end if
					end if
				
				end if
				
			end if ' pIdProduct<>""
			set rstemp=nothing
			call closeDb()
		end if

        '*************************************************************************************************
        ' START: GET CROSS SELL
        '*************************************************************************************************
        if (cint(pcCartArray(f,27))=-1) then    
            for t=1 to pcCartIndex
                if (pcCartArray(t,27) = f) AND (pcCartArray(t,12) = "0") then  '//  Bundled Child
                    pcCartArray(t,2) = pcCartArray(f,2)
                end if
            next
        end if
        '*************************************************************************************************
        ' END: GET CROSS SELL
        '*************************************************************************************************

	end if 'End of Finalized Quotes checking
	end if
next

call opendb()
Dim strCheckOver
strCheckOver = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)
If Len(Trim(strCheckOver))>0 then
   session("pcErrStrPrdDesc") =  strCheckOver
   session("pcErrIntStock") = pStock
   response.redirect "msg.asp?message=204"
End if

ppcCartIndex=pcCartIndex%>
<!--#include file="inc-UpdPrdQtyDiscounts.asp"-->
<%
session("pcCartSession")=pcCartArray

set f=nothing
set pcCartArray=nothing
set spcCartIndex=nothing
set pNewQuantity=nothing
set pcCartIndex=nothing

call clearLanguage()

'Calculate Product Promotions - START
%>
<!--#include file="inc_CalPromotions.asp"-->
<%
call closedb()
'Calculate Product Promotions - END
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing

response.redirect "viewCart.asp"
%>