<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "Custquotesview.asp"

'Thumbnail size: change the value below to change the size
Dim pcIntSmImgWidth
pcIntSmImgWidth = 25


' This page displays wishlist data and custom quote data.
'
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="CustLIv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->  
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/bto_language.asp" -->
<!--#include file="../includes/currencyformatinc.asp" --> 
<!--#include file="../includes/rc4.asp" -->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<%IF scBTO=1 THEN%>
<!--#include file="chkQuoteInfo.asp" -->
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="pcReCalQuotePricingCats.asp" -->
<%END IF%>
<!--#include file="../includes/adovbs.inc"-->
<%
Response.Buffer = True

Dim conntemp, query, rstemp 
Dim iPageSize,iPageCurrent, iPageCount,pCnt

iPageSize=5
iPageCurrent=getUserInput(request("iPageCurrent"),0)
if iPageCurrent="" then
	iPageCurrent=1
end if
if not IsNumeric(iPageCurrent) then
	response.redirect "CustPref.asp"
end if


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
call openDb()

pcv_HaveItems=0

pidcustomer=session("idcustomer")

if request.querystring("action")="del" then
	query="Delete FROM wishlist WHERE idCustomer="&pIdCustomer
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
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
	set rstemp = nothing
	call closedb()
	response.redirect "Custquotesview.asp?msg=" & Server.URLEncode(dictLanguage.Item(Session("language")&"_Custwlview_17"))
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Page On-Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

query="SELECT DISTINCT wishList.qsubmit,wishList.qdate,wishList.idquote,wishList.discountcode,configWishlistSessions.idconfigWishlistSession, configWishlistSessions.idproduct, configWishlistSessions.fPrice,configWishlistSessions.dPrice, products.description, products.serviceSpec, products.price, products.bToBPrice, products.SKU, products.stock, products.active, products.noprices,products.weight,products.noStock,products.smallImageUrl FROM (wishList INNER JOIN configWishlistSessions ON wishList.idconfigWishlistSession=configWishlistSessions.idconfigWishlistSession) INNER JOIN products ON configWishlistSessions.idproduct=products.idProduct WHERE (((wishList.idCustomer)="&pIdCustomer&") AND ((wishList.idconfigwishlistsession)<>0)) ORDER BY wishList.idquote DESC;"
Set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

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

if not rstemp.eof then
	iPageCount=rstemp.PageCount

	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
	rstemp.AbsolutePage=iPageCurrent
	pCnt=0
	do while not rstemp.eof and pCnt<iPageSize
		pCnt=pCnt+1
		QSubmit=rstemp("qsubmit")
		if IsNull(QSubmit) or QSubmit="" then
			QSubmit=0
		end if
		QIDQuote=rstemp("IDQuote")
		QIDConfig=rstemp("idconfigWishlistSession")
		QIDProduct=rstemp("idproduct")
		if QSubmit<1 then
			call updQuoteInfo(QIDQuote,QIDProduct,QIDConfig)
			if NOT isNULL(session("customerCategory")) and session("customerCategory")<>"" and session("customerCategory")<>0 then
				call updPricingCats(QIDQuote,QIDProduct,QIDConfig)
			end if
		end if
		rstemp.MoveNext
	loop
	
	set rstemp=nothing
end if

QuotesTotal=0

query="SELECT DISTINCT wishList.qsubmit,wishList.qdate,wishList.idquote,wishList.discountcode,configWishlistSessions.idconfigWishlistSession, configWishlistSessions.idproduct, configWishlistSessions.pcconf_Quantity, configWishlistSessions.fPrice,configWishlistSessions.dPrice, products.description, products.serviceSpec, products.price, products.bToBPrice, products.SKU, products.stock, products.active, products.noprices,products.weight,products.noStock,products.pcProd_BackOrder,products.smallImageUrl FROM (wishList INNER JOIN configWishlistSessions ON wishList.idconfigWishlistSession=configWishlistSessions.idconfigWishlistSession) INNER JOIN products ON configWishlistSessions.idproduct=products.idProduct WHERE (((wishList.idCustomer)="&pIdCustomer&")  AND ((wishList.idconfigwishlistsession)<>0)) ORDER BY wishList.idquote DESC;"
Set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

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

%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td>
			<% 
			msg=getUserInput(request.querystring("msg"),0)
			if msg="1" then
				msg=bto_dictLanguage.Item(Session("language")&"_Custquotesview_11")
			end if
			if msg<>"" then %>
				<div class="pcSuccessMessage">
					<%=msg%>
				</div>
			<%end if %>

	<%IF scBTO=1 THEN%>
	<%if not rstemp.eof then
	pcv_HaveItems=1%>
	<tr>
		<td> 
			<h1>
			<%
				if session("pcStrCustName") <> "" then
					response.write(session("pcStrCustName") & " - " &  bto_dictLanguage.Item(Session("language")&"_Custquotesview_1"))
					else
					response.write(bto_dictLanguage.Item(Session("language")&"_Custquotesview_1"))
				end if					
			%>
			</h1>
		</td>
	</tr>
	<tr>
		<td>
			<table class="pcShowContent">
				<tr> 
					<th width="5%" nowrap><%response.write bto_dictLanguage.Item(Session("language")&"_Custquotesview_4")%>
 #</th>
 					<th width="70%"><%response.write bto_dictLanguage.Item(Session("language")&"_Custquotesview_3")%></th>
					<th><%response.write dictLanguage.Item(Session("language")&"_Custwlview_10")%></th>
					<th>&nbsp;</th>
				</tr>
				<% 
				wishListTotal=Cint(0)
				iPageCount=rstemp.PageCount

				If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
				If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
				rstemp.AbsolutePage=iPageCurrent
				
				pCnt=0
				do while not rstemp.eof and pCnt<iPageSize
					pCnt=pCnt+1
					qsubmit=rstemp("qsubmit")
					qdate=rstemp("qdate")
					idquote=rstemp("idquote")
					pdiscountcode=rstemp("discountcode")
					if pdiscountcode="0" then
						pdiscountcode=""
					end if
					QIDProduct=rstemp("idproduct")
					pQty=rstemp("pcconf_Quantity")
					fPrice=rstemp("fPrice")
					dPrice=rstemp("dPrice")
					pBtoBPrice=rstemp("bToBPrice")
					pPrice=rstemp("price")
					pserviceSpec=rstemp("serviceSpec")
					pnoprices=rstemp("noprices")
					if pnoprices<>"" then
					else
					pnoprices=0
					end if
					if qsubmit=3 then
						pnoprices=0
					end if
					
					pweight=clng(rstemp("weight"))
					pNoStock=rstemp("noStock")
					pcv_BackOrder=rstemp("pcProd_BackOrder")
					pcv_smallImage=rstemp("smallImageUrl")
									
					if pserviceSpec=true then
						query="SELECT categories.categoryDesc, products.description, configSpec_products.configProductCategory, configSpec_products.price, configSpec_products.Wprice, categories_products.idCategory, categories_products.idProduct, products.weight FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&rstemp("idproduct")&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct]) AND ((configSpec_products.cdefault)<>0)) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
						set rsSSObj=conntemp.execute(query)
						if err.number<>0 then
							'//Logs error to the database
							call LogErrorToDatabase()
							'//clear any objects
							set rsSSObj=nothing
							'//close any connections
							call closedb()
							'//redirect to error page
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
															
						if NOT rsSSobj.eof then 
							Dim iAddDefaultPrice,	iAddDefaultWPrice
							iAddDefaultPrice=Cdbl(0)
							iAddDefaultWPrice=Cdbl(0)
							do until rsSSobj.eof
								iAddDefaultPrice=Cdbl(iAddDefaultPrice+rsSSobj("price"))
								iAddDefaultWPrice=Cdbl(iAddDefaultWPrice+rsSSobj("Wprice")) 
							rsSSobj.moveNext
							loop
							set rsSSobj=nothing
							pPrice=Cdbl(pPrice+iAddDefaultPrice)
							pBtoBPrice=Cdbl(pBtoBPrice+iAddDefaultWPrice)
						end if
					end if 
															 
					pShowPrice=0
					if pBtoBPrice>"0" and session("customerType")="1" then
						wishListTotal=wishListTotal+pBtoBPrice
						pShowPrice=pBtoBPrice
					else
						wishListTotal=wishListTotal+pPrice
						pShowPrice=pPrice
					end if
					%>
					<tr> 
						<td valign="top" nowrap>
							<%response.write rstemp("idquote")%>
						</td>
						
						<%
						' Check if thumbnail exists
							if pcv_smallImage = "" or pcv_smallImage = "no_image.gif" then
								pcv_smallImage = "hide"
							end if
						%>
						
						<td valign="top"> 
							<% if pcv_smallImage <> "hide" then %>
							<img src="catalog/<%=pcv_smallImage%>" width="<%=pcIntSmImgWidth%>" align="middle" style="border:none; padding-top: 2px; padding-bottom: 5px; padding-right: 4px;">
							<% end if
							response.write rstemp("description") &" ("& rstemp("sku") & ")"
							if clng(pQty)>1 then
							response.write " - " & bto_dictLanguage.Item(Session("language")&"_printableQuote_6") &":" & pQty
							end if%>
							<% if rstemp("active")=0 then %>
								<br>
								<%response.write bto_dictLanguage.Item(Session("language")&"_Custquotesview_6")%>
								<br>
							<% else
								pStock=rstemp("stock")
								if (scShowStockLmt=-1 AND pNoStock=0 AND int(pStock)<1 AND pserviceSpec=false AND pcv_BackOrder=0) OR (pserviceSpec=true AND scShowStockLmt=-1 AND iBTOOutofstockpurchase=-1 AND int(pStock)<1 AND pNoStock=0 AND pcv_BackOrder=0) then %>
									<br>
									<%response.write dictLanguage.Item(Session("language")&"_viewPrd_7")%>
									<br>
								<%end if
							end if %>
							<br>
							<%'discounts by categories
							CatDiscTotal=0
							
							query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
							set rs1=conntemp.execute(query)
							if err.number<>0 then
								'//Logs error to the database
								call LogErrorToDatabase()
								'//clear any objects
								set rs1=nothing
								'//close any connections
								call closedb()
								'//redirect to error page
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							CatSubDiscount=0
							
							Do While (not rs1.eof) and (CatSubDiscount=0)
							 CatSubQty=0
							 CatSubTotal=0
							 CatSubDiscount=0
							
								query="select idproduct from categories_products where idcategory=" & rs1("IDCat") & " and idproduct=" & QIDProduct
								set rs=connTemp.execute(query)
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
								
								if not rs.eof then
									CatSubQty=CatSubQty+pQty
									CatSubTotal=CatSubTotal+fPrice
								end if
							
							if CatSubQty>0 then
							
							query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rs1("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
							set rs2=conntemp.execute(query)
							if err.number<>0 then
								'//Logs error to the database
								call LogErrorToDatabase()
								'//clear any objects
								set rs2=nothing
								'//close any connections
								call closedb()
								'//redirect to error page
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							if not rs2.eof then
							
								' there are quantity discounts defined for that quantity 
								pDiscountPerUnit=rs2("pcCD_discountPerUnit")
								pDiscountPerWUnit=rs2("pcCD_discountPerWUnit")
								pPercentage=rs2("pcCD_percentage")
								pbaseproductonly=rs2("pcCD_baseproductonly")
							
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
							
									
							end if
							end if
							
								CatDiscTotal=CatDiscTotal+CatSubDiscount
								rs1.MoveNext
							loop
							
							'// Round the Category Discount to two decimals
							if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
								CatDiscTotal = Round(CatDiscTotal,2)
							end if
				
							if pnoprices<2 then
								if CatDiscTotal>0 then%>
									<%=dictLanguage.Item(Session("language")&"_catdisc_2")%> -<%response.write scCurSign & money(CatDiscTotal)%>
									<br>
								<%end if
							end if%>
					<% if pDiscountCode<>"" then %>
					<!--#include file="checkDiscount.asp"-->
					<%
					discountcheck=0
				
					if pDiscountError="" then
						discountcheck=1 %>
						<i><%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_15")%></i><br>
						<%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_12")%>: <%=pDiscountCode%><br>
						<%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_13")%>: <%=pDiscountDesc%><br>
						<%if pFreeShip<>"" then%>
						<%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_14")%>: <%=pFreeShip%><br>
						<%else%>
						<%if pnoprices<2 then 'Does not hide prices
							if ppercentageToDiscount="0" then %>
								<%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_16")%>: - <%=scCurSign & money(pPriceToDiscount)%><br>
							<% else %>
								<%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_16")%>: - <%=ppercentageToDiscount %>%<br>
							<% end if
						end if %>
						<%end if%>
					<% else %>
					<%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_12")%>: <%=pDiscountCode%></font><br>
					<%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_17")%>: <%=pDiscountError%></font><br>
					<% end if 
				end if %>
				<div class="pcSmallText" style="padding: 5px 0 5px 0;"><a href="<%if qsubmit<2 then%>printableQuote.asp<%else%>printableEditedQuote.asp<%end if%>?w=<%=pWeight%>&idconf=<%=rstemp("idconfigWishlistSession")%>" target="_blank"><%=bto_dictLanguage.Item(Session("language")&"_Custquotesview_7")%></a></div>
					<% if cint(qsubmit)>0 then%>
						<%response.write bto_dictLanguage.Item(Session("language")&"_Custquotesview_8")%>&nbsp;<%=scCompanyName%>&nbsp;<%response.write bto_dictLanguage.Item(Session("language")&"_Custquotesview_9")%>&nbsp;<%=month(qdate)%>/<%=day(qdate)%>/<%=year(qdate)%>
					<%end if%>

					<% pIdConfigSession=trim(rstemp("idconfigWishlistSession"))
					if pIdConfigSession<>"0" then
					'BTO Items
								query="SELECT xfdetails, stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configWishlistSessions WHERE idconfigWishlistSession=" & pIdConfigSession
								set rsConfigObj=server.CreateObject("ADODB.RecordSet")
								set rsConfigObj=connTemp.execute(query)
								if err.number<>0 then
									'//Logs error to the database
									call LogErrorToDatabase()
									'//clear any objects
									set rsConfigObj=nothing
									'//close any connections
									call closedb()
									'//redirect to error page
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
			
								xstr=rsConfigObj("xfdetails")
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
								if (stringProducts<>"") and (stringProducts<>"na") then%>
								<table class="pcShowBTOconfiguration" width="100%">
								<tr> 
									<td colspan="2"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_4")%></td>
								</tr>
								<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
								set rsConfigObj=server.CreateObject("ADODB.RecordSet")
								set rsConfigObj=connTemp.execute(query)
								if rsConfigObj.EOF then 
								%>
									<tr>
										<td colspan="2"><%=dictLanguage.Item(Session("language")&"_Custwlview_13")%></td>
									</tr>
								<%
								else
									
								if err.number<>0 then
									'//Logs error to the database
									call LogErrorToDatabase()
									'//clear any objects
									set rsConfigObj=nothing
									'//close any connections
									call closedb()
									'//redirect to error page
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								pcategoryDesc=rsConfigObj("categoryDesc")
								pdescription=rsConfigObj("description")
								psku=rsConfigObj("sku")
								pItemWeight=rsConfigObj("weight")
								set rsConfigObj=nothing

								query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & QIDProduct 
								set rs=server.CreateObject("ADODB.RecordSet") 
								set rs=conntemp.execute(query)
								if NOT rs.eof then				
									btDisplayQF=rs("displayQF")
								end if
								set rs=nothing
											
								query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
								set rsQ=connTemp.execute(query)
								tmpMinQty=1
								if not rsQ.eof then
									tmpMinQty=rsQ("pcprod_minimumqty")
									if IsNull(tmpMinQty) or tmpMinQty="" then
										tmpMinQty=1
									else
										if tmpMinQty="0" then
											tmpMinQty=1
										end if
									end if
								end if
								set rsQ=nothing
								tmpDefault=0
								query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & QIDProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
								set rsQ=connTemp.execute(query)
								if not rsQ.eof then
									tmpDefault=rsQ("cdefault")
									if IsNull(tmpDefault) or tmpDefault="" then
										tmpDefault=0
									else
										if tmpDefault<>"0" then
										 	tmpDefault=1
										end if
									end if
								end if
								set rsQ=nothing %>
								<tr> 
									<td width="30%"><%=pcategoryDesc%>:</td>
									<% if NOT isNumeric(ArrQuantity(i)) then
											pIntQty=1
										else
											pIntQty=ArrQuantity(i)
										end if %>
									<td width="70%"><%=psku%> - <%=pdescription%>
									<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%> - <%response.write dictLanguage.Item(Session("language")&"_custOrdInvoice_18")%>: <%=ArrQuantity(i)%><%end if%>
									</td>
								</tr>
							<%
							end if
						next
						set rsConfigObj=nothing
						%>
						</table>
					<%end if%>
					<% 'BTO Additional Charges
						query="SELECT stringCProducts,stringCValues,stringCCategories FROM configWishlistSessions WHERE idconfigWishlistSession=" & pIdConfigSession
						set rsConfigObj=server.CreateObject("ADODB.RecordSet")
						set rsConfigObj=connTemp.execute(query)
						if err.number<>0 then
							'//Logs error to the database
							call LogErrorToDatabase()
							'//clear any objects
							set rsConfigObj=nothing
							'//close any connections
							call closedb()
							'//redirect to error page
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						
						stringCProducts=rsConfigObj("stringCProducts")
						stringCValues=rsConfigObj("stringCValues")
						stringCCategories=rsConfigObj("stringCCategories")
						ArrCProduct=Split(stringCProducts, ",")
						ArrCValue=Split(stringCValues, ",")
						ArrCCategory=Split(stringCCategories, ",")
						
						if ArrCProduct(0)<>"na" then
						%>
						<table class="pcShowBTOconfiguration" width="100%">
							<tr> 
								<td colspan="2"><%response.write bto_dictLanguage.Item(Session("language")&"_printableQuote_5")%></td>
							</tr>
		
											<% Charges=0
											for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
												query="SELECT categories.categoryDesc, products.description, products.sku, products.weight FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
												set rsConfigObj=server.CreateObject("ADODB.RecordSet")
												set rsConfigObj=connTemp.execute(query)
												if err.number<>0 then
													'//Logs error to the database
													call LogErrorToDatabase()
													'//clear any objects
													set rsConfigObj=nothing
													'//close any connections
													call closedb()
													'//redirect to error page
													response.redirect "techErr.asp?err="&pcStrCustRefID
												end if
												
												pcategoryDesc=rsConfigObj("categoryDesc")
												pdescription=rsConfigObj("description")
												psku=rsConfigObj("sku")
												pItemWeight=rsConfigObj("weight")
												intTotalWeight=intTotalWeight+int(pItemWeight)
												if (CDbl(ArrCValue(i))>0)then
													Charges=Charges+cdbl(ArrCValue(i))
												end if %>
												<tr> 
													<td width="30%"><%=pcategoryDesc%>:</td>
													<td width="70%"><%=psku%> - <%=pdescription%></td>
												</tr>
												<% set rsConfigObj=nothing
											next
											pRowPrice=pRowPrice+Cdbl(Charges)%>
										</table>
								<% end if
								'BTO Additional Charges %>
					<% end if 'Have idconfigWishlistSession%>
							
					</td>
					<td valign="top" nowrap> 
					<%
					if pnoprices<2 then
						if discountcheck=1 then
							if pDiscountError="" then    
								discountTotal=Cdbl(0)
								if pPriceToDiscount>0 or ppercentageToDiscount>0 then 
									discountTotal=pPriceToDiscount + (ppercentageToDiscount*(fPrice)/100)
								end if
								dim pSubTotal
								pSubTotal=fPrice - discountTotal
								if pSubTotal<0 then
									pSubTotal=0
								end if
							end if
							if CatDiscTotal>0 then
							pSubTotal=pSubTotal - CatDiscTotal
							if pSubTotal<0 then
							pSubTotal=0
							end if
							end if
							response.write scCurSign & money(pSubTotal)
						else
							pSubTotal=fPrice
							if CatDiscTotal>0 then
							pSubTotal=pSubTotal - Round(CatDiscTotal,2)
							if pSubTotal<0 then
							pSubTotal=0
							end if
							end if
							response.write scCurSign & money(pSubTotal)
						end if
						QuotesTotal=QuotesTotal+pSubTotal
					end if%>
					</td>
					<td valign="top" nowrap>
					<div align="right">
					<% 
					if (iBTOQuoteSubmitOnly=0) and (pnoprices=0) and cint(qSubmit)<1 then
						if rstemp("active")<>"0" then %>
							<a href="Reconfigure.asp?price=<%=rstemp("fPrice")%>&idproduct=<%=rstemp("idproduct")%>&idconf=<%=rstemp("idconfigWishlistSession")%>&act=placeOrder"><img src="<%=rslayout("revorder")%>" border="0" vspace="2"></a><br>
						<% end if
					end if %>
					<%if (cint(qSubmit)=0) and (pnoprices>0) then%>
						<a href="Reconfigure.asp?price=<%=rstemp("fPrice")%>&idproduct=<%=rstemp("idproduct")%>&idconf=<%=rstemp("idconfigWishlistSession")%>"><img src="<%=rslayout("reconfigure")%>" border="0" vspace="2" alt="Edit quote"></a><br>
					<%end if%>
					<% if (cint(qSubmit)=0) and (iBTOQuoteSubmit=1) then%>
						<a href="mySubmitQuote.asp?idcustomer=<%=session("idcustomer")%>&idproduct=<%=rstemp("idproduct")%>&idconf=<%=rstemp("idconfigWishlistSession")%>"><img src="<%=rslayout("submitquote")%>" border="0" vspace="2"></a><br>
					<% end if %>
					<%if cint(qSubmit)=3 then%>
						<a href="AddFQuoteToCart.asp?idconf=<%=rstemp("idconfigWishlistSession")%>"><img src="<%=rslayout("pcLO_placeOrder")%>" border="0" vspace="2"></a><br>
					<%end if%>
					<a href="CustquoteRmv.asp?idconf=<%=rstemp("idconfigWishlistSession")%>&iPageCurrent=<%=iPageCurrent%>"><img src="<%=rslayout("remove")%>" border="0" vspace="2"></a></div></td>
					</tr>
					<tr>
						<td colspan="4"></td>
					</tr>
					<!-- START - Custom Input Fields -->
					<% 'if there are custom input fields, show them here
					Dim xstr
					if trim(xstr)<>"" then 
					%>
					<tr>
						<td>&nbsp;</td>
						<td colspan="3">
							<p>
							<%
							if xstr<>"" then
								xarray=split(xstr,"||")
							end if	
							For x=lbound(xarray) to ubound(xarray)
								if xarray(x)<>"" then
									temparray=split(xarray(x),"|")
									pxfield=temparray(0)
									pxvalue=temparray(1)
									if pxfield<>"0" then
										'select from the database more info
										mySQL= "SELECT xfields.xfield FROM xfields WHERE idxfield="&pxfield
										set rsXFieldObj=server.CreateObject("ADODB.RecordSet")
										set rsXFieldObj=connTemp.execute(mySQL)									
										if err.number <> 0 then
											call closeDb()
											set rsXFieldObj=nothing
											response.redirect "techErr.asp?error="& Server.Urlencode("Error in viewPrd: "&err.description) 
										end if 										
										if NOT rsXFieldObj.eof then
											response.write "<div>" & rsXFieldObj("xfield") & ": " & pxvalue & "</div>"
										end if
										set rsXFieldObj=nothing
									end if
								end if
							Next
							%>
							</p>
						</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<%
					End if '// if trim(pcCartArray(f,21))<>"" then 
					%>						            
					<!-- END - Custom Input Fields -->				
					<tr>
						<td colspan="4"><hr></td>
					</tr>
				<%
				rstemp.movenext
				loop
				%>
				<tr>
					<td colspan="4">
						<%
						iRecSize=10

						'*******************************
						' START Page Navigation
						'*******************************

						If iPageCount>1 then %>
							<div class="pcPageNav">
							<%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
							&nbsp;-&nbsp;
						    <% if iPageCount>iRecSize then %>
								<% if cint(iPageCurrent)>iRecSize then %>
									<a href="Custquotesview.asp?iPageCurrent=1"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_1")%></a>&nbsp;
					        	<% end if %>
								<% if cint(iPageCurrent)>1 then
	            					if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
	                					iPagePrev=cint(iPageCurrent)-1
	            					else
	                					iPagePrev=iRecSize
	            					end if %>
	            					<a href="Custquotesview.asp?iPageCurrent=<%=cint(iPageCurrent)-iPagePrev%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_2")%>&nbsp;<%=iPagePrev%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
								<% end if
								if cint(iPageCurrent)+1>1 then
									intPageNumber=cint(iPageCurrent)
								else
									intPageNumber=1
								end if
							else
								intPageNumber=1
							end if

							if (cint(iPageCount)-cint(iPageCurrent))<iRecSize then
								iPageNext=cint(iPageCount)-cint(iPageCurrent)
							else
								iPageNext=iRecSize
							end if

							For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
								If Cint(pageNumber)=Cint(iPageCurrent) Then %>
									<strong><%=pageNumber%></strong> 
								<% Else %>
		      						<a href="Custquotesview.asp?iPageCurrent=<%=pageNumber%>"><%=pageNumber%></a>
								<% End If 
							Next
	
							if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
							else
								if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
									<a href="Custquotesview.asp?iPageCurrent=<%=cint(intPageNumber)+iPageNext%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_4")%>&nbsp;<%=iPageNext%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
								<% end if
    
								if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
						    		&nbsp;<a href="Custquotesview.asp?iPageCurrent=<%=cint(iPageCount)%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_5")%></a>
						    	<% end if 
							end if 

						end if

						'*******************************
						' END Page Navigation
						'*******************************
						%>
					</td>
				</tr>
			</table>
			<%end if%>
			<%END IF 'ProductCart BTO%>
			
			<%
			query="SELECT products.idProduct, products.description, products.serviceSpec, products.price, products.sku, products.bToBPrice, products.stock, products.noStock, products.active, products.smallImageUrl, wishlist.pcwishlist_OptionsArray, wishlist.IDQuote,products.pcProd_BackOrder FROM wishlist, products WHERE (((products.idProduct)=[wishlist].[idproduct]) AND ((wishlist.idCustomer)="&pIdCustomer&") AND ((products.serviceSpec)=0));"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
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
			
			if not rstemp.eof then
			pcv_HaveItems=1%>
			<h1>
				<%
				if session("pcStrCustName") <> "" then
					response.write(session("pcStrCustName") & " - " &  dictLanguage.Item(Session("language")&"_Custwlview_1"))
					else
					response.write(dictLanguage.Item(Session("language")&"_Custwlview_1"))
				end if
				%>
			</h1>
			<table class="pcShowContent">
				<tr> 
					<th width="5%" nowrap><%response.write dictLanguage.Item(Session("language")&"_Custwlview_8")%></th>
					<th width="70%" nowrap><%response.write dictLanguage.Item(Session("language")&"_Custwlview_9")%></th>
					<th align="center"><%response.write dictLanguage.Item(Session("language")&"_Custwlview_10")%></th>
					<th>&nbsp;</th>
				</tr>
				<% wishListTotal=Cint(0)
				do while not rstemp.eof 
					pIdProduct=rstemp("idProduct")
					pDescription=rstemp("description")
					pserviceSpec=rstemp("serviceSpec")
					pPrice=rstemp("price")
					pSku=rstemp("sku")
					pBtoBPrice=rstemp("bToBPrice")
					if pBtoBPrice=0 then
						pBtoBPrice=pPrice
					end if
					pStock=rstemp("stock")
					pNoStock=rstemp("noStock")
					pActive=rstemp("active")
					pcv_strSelectedOptions=rstemp("pcwishlist_OptionsArray")
					pcv_strIDQuote=rstemp("IDQuote")
					pcv_BackOrder=rstemp("pcProd_BackOrder")
					pcv_smallImage=rstemp("smallImageUrl")
					
					dblpcCC_Price=0
					
					'Check if this customer is logged in with a customer category
					if session("customerCategory")<>0 then
						query="SELECT idCC_Price, pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
						if NOT rs.eof then
							idCC_Price=rs("idCC_Price")
							dblpcCC_Price=rs("pcCC_Price")
							dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
							if dblpcCC_Price>0 then
								strcustomerCategory="YES"
							else
								strcustomerCategory="NO"
							end if
						else
							strcustomerCategory="NO"
						end if
						set rs=nothing
					end if
					
					if session("customerCategoryType")="ATB" then
						if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
							pPrice=pPrice-(pcf_Round(pPrice*(cdbl(session("ATBPercentage"))/100),2))
						end if
						if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
							pBtoBPrice=pBtoBPrice-(pcf_Round(pBtoBPrice*(cdbl(session("ATBPercentage"))/100),2))
							pPrice=pBtoBPrice
						end if						
					end if
					
					if pBtoBPrice>"0" and session("customerType")=1 then
						pPrice=pBtoBPrice
					end if
					
					if strcustomerCategory="YES" then
						pPrice=dblpcCC_Price
					end if

				'*************************************************************************************************
				' START: GET OPTIONS
				'*************************************************************************************************
				Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc, pcv_strSelectedOptions
				Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
				Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START:  Get the Options for the item
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				if pcv_strSelectedOptions<>"" then
				pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))
				
				pcv_strOptionsArray = ""
				pcv_strOptionsPriceArray = ""
				pcv_strOptionsPriceArrayCur = ""
				pcv_strOptionsPriceTotal = 0
				xOptionsArrayCount = 0
				
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
					
					if Not rs.eof then 
						
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
				end if	
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  Get the Options for the item
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
			
							
				'*************************************************************************************************
				' END: GET OPTIONS
				'*************************************************************************************************
					
				pShowPrice=0
				if pcv_strOptionsPriceTotal = "" then
					pcv_strOptionsPriceTotal = 0
				end if
				
				wishListTotal = wishListTotal+pPrice+pcv_strOptionsPriceTotal
				pShowPrice = pPrice + pcv_strOptionsPriceTotal
				
				' Check if thumbnail exists
					if pcv_smallImage = "" or pcv_smallImage = "no_image.gif" then
						pcv_smallImage = "hide"
					end if
				%>

				<tr> 
					<td valign="top" nowrap> 
						<%response.write pSku%>
					</td>
					<td valign="top">
						<% if pcv_smallImage <> "hide" then %>
						<img src="catalog/<%=pcv_smallImage%>" width="<%=pcIntSmImgWidth%>" align="middle" style="border:none; padding-top: 2px; padding-bottom: 5px; padding-right: 4px;">
						<% end if	%>
						<% if pActive<>"0" then %>
							<a href="viewPrd.asp?idproduct=<%response.write pIdProduct%>&idOptionArray=<%=pcv_strSelectedOptions%>"><%response.write pDescription%></a>
						<% else %>
							<%response.write pDescription%>
						<% end if %>
						
						<% if pActive=0 then %>
							<br>
							<%response.write dictLanguage.Item(Session("language")&"_Custwlview_13")%>
							<br>
						<% else %>
							<% if scShowStockLmt=-1 AND pStock<1 AND pNoStock=0 AND pcv_BackOrder=0 then %>
							<br>
							<%response.write dictLanguage.Item(Session("language")&"_viewPrd_7")%>
							<br>
							<%
							end if
						end if 
						%>

					</td>
					<td align="right" valign="top" nowrap>
						<%response.write scCurSign & money(pShowPrice)%>
					</td>
					<td valign="top">
            <a href="CustwlRmv.asp?IDQuote=<%=pcv_strIDQuote%>"><img src="<%=rslayout("remove")%>"></a>
					</td>
				</tr>
				<tr>
					<td> </td>
					<td colspan="2">					
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: SHOW PRODUCT OPTIONS
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					Dim pcArray_strOptionsPrice, pcArray_strOptions, pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice, tAprice
					
					if len(pcv_strOptionsArray)>0 then 
					%>
					
					
					<%'response.write dictLanguage.Item(Session("language")&"_Custwlview_15")%>
					
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
					
						<%
						'#####################
						' START LOOP
						'#####################	
						'// Generate Our Local Arrays from our Stored Arrays
						
						' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
						pcArray_strSelectedOptions = ""					
						pcArray_strSelectedOptions = Split(trim(pcv_strSelectedOptions),chr(124))
						
						' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
						pcArray_strOptionsPrice = ""
						pcArray_strOptionsPrice = Split(trim(pcv_strOptionsPriceArray),chr(124))
						
						' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
						pcArray_strOptions = ""
						pcArray_strOptions = Split(trim(pcv_strOptionsArray),chr(124))
						
						' Get Our Loop Size
						pcv_intOptionLoopSize = 0
						pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
						
						' Start in Position One
						pcv_intOptionLoopCounter = 0
						
						' Display Our Options
						For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
						%>
						<tr>
						<td width="75%"><p><%= pcArray_strOptions(pcv_intOptionLoopCounter)%></p></td>
						
							<td align="right" width="25%">									
							<% 
							tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
							
							if tempPrice="" or tempPrice=0 then
								response.write "&nbsp;"
							else 
								response.write scCurSign&money(tempPrice)
							end if 
							%>			
							
							</td>
						</tr>
						<%
						Next
						'#####################
						' END LOOP
						'#####################
						%>
					
					</table>
					
					<% 
					End if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: SHOW PRODUCT OPTIONS
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					%>					
					</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="4"><hr></td>
				</tr>
				<%
				rstemp.movenext
				loop
				%>
			<%if iPageCount=1 then%>
			<tr>
				<td align="right" colspan="2"><%response.write dictLanguage.Item(Session("language")&"_Custwlview_5")%></td>
				<td align="left" nowrap>					
					<%response.write " " & scCurSign & money(wishListTotal+QuotesTotal)%>
				</td>
				<td>&nbsp;</td>
			</tr>
			<%end if%>
			<%if pcv_HaveItems>0 then%>
			<%if iPageCount=1 then%>
			<tr>
				<td colspan="4"><hr></td>
			</tr>
			<%end if%>
			<tr>
				<td align="center" colspan="4">
					<%
					Dim pcv_allowPurchase
					if (scorderlevel="1" AND session("customerType")=1) or scorderlevel="0" then
						pcv_allowPurchase = 1
					else
						pcv_allowPurchase = 0
					end if
					
					if pcv_allowPurchase = 1 then
					%>
						<a href="addsavedprdstocart.asp"><%response.write dictLanguage.Item(Session("language")&"_Custwlview_14")%></a>&nbsp;|&nbsp;
					<%
					end if
					%>
					<a href="javascript:if (confirm('<%response.write dictLanguage.Item(Session("language")&"_Custwlview_18")%>')) location='Custquotesview.asp?action=del&iPageCurrent=<%=iPageCurrent%>'"><%response.write dictLanguage.Item(Session("language")&"_Custwlview_16")%></a>
					</td>
			</tr>
			<%end if%>
			</table>
			<%end if%>
			
			<table class="pcShowContent">
			<%if pcv_HaveItems=0 then%>
				<tr>
					<td><%response.write dictLanguage.Item(Session("language")&"_Custwlview_2")%></td>
				</tr>
			<%end if%>
				<tr>
					<td class="pcSpacer"></td>
				</tr>
				<tr>
					<td align="center">					
						<a href="default.asp"><img src="<%=RSlayout("continueshop")%>" border="0" alt=""></a>
						<%
						if pcv_allowPurchase = 1 then
						%>
						&nbsp;
						<a href="viewCart.asp"><img src="<%=RSlayout("viewcartbtn")%>" border="0" alt=""></a>
						<%
						end if
						%>
						&nbsp;
						<a href="custpref.asp"><img src="<%=rslayout("back")%>" border="0" alt=""></a>
					</td>
				</tr>
			</table>
		 </td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->