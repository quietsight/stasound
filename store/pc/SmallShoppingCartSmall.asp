<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'Show shopping cart total
dim vcCartArr, vcCartIndex, v, scantidadCart7, vcTotal, vcItems

vcCartArr=Session("pcCartSession")
vcCartIndex=Session("pcCartIndex")
vcTotal=Cint(0) 'calculates the cart total
vcItems=Cint(0)	'counts items in your cart
vcPrice=Cint(0) 'calculates the cart cross sell product
vcpPrice=Cint(0) 'calculates the cart cross sell parent product

opcOrderTotal=0

Dim sscProList(100,5)

for v=1 to vcCartIndex
	sscProList(v,0)=vcCartArr(v,0)
	sscProList(v,1)=vcCartArr(v,10)
	sscProList(v,3)=vcCartArr(v,2)
	sscProList(v,4)=0
					
	if InStr(Cstr(10/3),",")>0 then
		if Instr(vcCartArr(v,17),".")>0 then
			if IsNumeric(vcCartArr(v,17)) then
				vcCartArr(v,17)=replace(vcCartArr(v,17),".",",")
			end if
		end if
	else
		if Instr(vcCartArr(v,17),",")>0 then
			if IsNumeric(vcCartArr(v,17)) then
				vcCartArr(v,17)=replace(vcCartArr(v,17),",",".")
			end if
		end if
	end if

	if vcCartArr(v,10)=0 then
		if vcCartArr(v,2)="" then
			vcCartArr(v,2)=1
		end if
		if vcCartArr(v,17)="" then
			vcCartArr(v,17)=0
		end if
		if vcCartArr(v,5)="" then
			vcCartArr(v,5)=0
		end if
		vcItems=vcItems + vcCartArr(v,2)
		vcPrice=Cint(0)
		vcPrice=vcPrice+(vcCartArr(v,2)*vcCartArr(v,17))
		vcPrice=vcPrice+(vcCartArr(v,2)*vcCartArr(v,5))
		if vcCartArr(v,30)<>"" then
		vcPrice=vcPrice-Abs(vcCartArr(v,30))
		end if
		if vcCartArr(v,31)<>"" then
		vcPrice=vcPrice+vcCartArr(v,31)
		end if
		vcPrice=vcPrice-vcCartArr(v,15)
		
		if trim(vcCartArr(v,27))="" then
			vcCartArr(v,27)=0
		end if
		if trim(vcCartArr(v,28))="" then
			vcCartArr(v,28)=0
		end if
		
		if (vcCartArr(v,27)>"0") AND (vcCartArr(v,28)>"0") then 
		    vp = cint(vcCartArr(v,27))
			vcpPrice=Cint(0)
			vcpPrice=vcpPrice+(vcCartArr(vp,2)*vcCartArr(vp,17))
			vcpPrice=vcpPrice+(vcCartArr(vp,2)*vcCartArr(vp,5))
			vcpPrice=vcpPrice-Abs(vcCartArr(vp,30))
			vcpPrice=vcpPrice+vcCartArr(vp,31)
			vcpPrice=vcpPrice-vcCartArr(vp,15)
			vcPrice = (cdbl(vcPrice)+cdbl(vcpPrice)) - ((cdbl(vcCartArr(v,28)) + cdbl(vcCartArr(vp,28)))*vcCartArr(v,2))
		end if		
		
		'SB S
		If vcCartArr(v,38)>"0" Then
			pSubscriptionIDb = (vcCartArr(v,38)) 
			Set conTempSC=Server.CreateObject("ADODB.Connection")
			conTempSC.Open scDSN
			query="SELECT SB_IsTrial, SB_TrialAmount FROM SB_Packages WHERE SB_PackageID=" & pSubscriptionIDb 
			set rsSub=server.CreateObject("ADODB.RecordSet")
			set rsSub=conTempSC.execute(query)
			if not rsSub.eof then
				pcv_intIsTrial=rsSub("SB_IsTrial")
				pcv_curTrialAmount = rsSub("SB_TrialAmount")
				if pcv_intIsTrial = "1" then
					vcPrice = (cdbl(pcv_curTrialAmount)*vcCartArr(v,2))
				end if	
			end if
			conTempSC.Close
			Set conTempSC=nothing
			set rsSub = nothing
		 End if
		'SB E

		sscProList(v,2)=vcPrice
		
		'// Don't Add to total if parent of a Bundle Cross Sell Product
		pcv_HaveBundles=0
		if vcCartArr(v,27)=-1 then
			for mc=1 to vcCartIndex
				if (vcCartArr(mc,27)<>"") AND (vcCartArr(mc,12)<>"") then
					if cint(vcCartArr(mc,27))=v AND cint(vcCartArr(mc,12))="0" then
						pcv_HaveBundles=1
						exit for
					end if
				end if
			next
		end if
		if (vcCartArr(v,27)>-1) OR (pcv_HaveBundles=0) then
			vcTotal=vcTotal+vcPrice
		end if		
	end if
	
next

' ------------------------------------------------------
' START - Calculate category-based quantity discounts
' ------------------------------------------------------
Set conTempSC=Server.CreateObject("ADODB.Connection")
conTempSC.Open scDSN

CatDiscTotal=0

query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
set rsSSCCatDis=server.CreateObject("ADODB.RecordSet")
set rsSSCCatDis=conTempSC.execute(query)
Do While not rsSSCCatDis.eof
	CatSubQty=0
	CatSubTotal=0
	CatSubDiscount=0

	for v=1 to vcCartIndex
		if (sscProList(v,1)=0) and (sscProList(v,4)=0) then
			if (vcCartArr(v,32)<>"") then
				pcv_tmpPPrd=split(vcCartArr(v,32),"$$")
				pcv_tmpID=pcv_tmpPPrd(ubound(pcv_tmpPPrd))
			else
				pcv_tmpID=sscProList(v,0)
			end if
			query="select idproduct from categories_products where idcategory=" & rsSSCCatDis("IDCat") & " and idproduct=" & pcv_tmpID
			set rsSSCProd=server.CreateObject("ADODB.RecordSet")
			set rsSSCProd=conTempSC.execute(query)
			if not rsSSCProd.eof then
				CatSubQty=CatSubQty+sscProList(v,3)
				CatSubTotal=CatSubTotal+sscProList(v,2)
				sscProList(v,4)=1
			end if
			set rsSSCProd=nothing
		end if
	Next

	if CatSubQty>0 then
		query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rsSSCCatDis("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
		set rsSSCDiscount=server.CreateObject("ADODB.RecordSet")
		set rsSSCDiscount=conTempSC.execute(query)
		if not rsSSCDiscount.eof then

			' there are quantity discounts defined for that quantity 
			pDiscountPerUnit=rsSSCDiscount("pcCD_discountPerUnit")
			pDiscountPerWUnit=rsSSCDiscount("pcCD_discountPerWUnit")
			pPercentage=rsSSCDiscount("pcCD_percentage")
			pbaseproductonly=rsSSCDiscount("pcCD_baseproductonly")
			set rsSSCDiscount=nothing
			
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
	rsSSCCatDis.MoveNext
loop
set rsSSCCatDis=nothing	

'// Round the Category Discount to two decimals
if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
	CatDiscTotal = Round(CatDiscTotal,2)
	vcTotal=vcTotal-CatDiscTotal
end if

'// Display Applied Product Promotions (if any)
TotalPromotions=0
if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
	TotalPromotions=pcf_GetPromoTotal(Session("pcPromoSession"),Session("pcPromoIndex"))
	vcTotal=vcTotal-TotalPromotions
end if

conTempSC.Close
Set conTempSC=nothing
' ------------------------------------------------------
' END - Calculate category-based quantity discounts
' ------------------------------------------------------



if vcItems > 0 then

	vcHaveGcsTest=0
	dim pcv_counterb
	for pcv_counterb=1 to vcCartIndex
		if vcCartArr(pcv_counterb,10)=0 then
			Set conTempSC=Server.CreateObject("ADODB.Connection")
			conTempSC.Open scDSN
			query="select pcprod_Gc from Products where idproduct=" & vcCartArr(pcv_counterb,0) & " AND pcprod_Gc=1"
			set rsGcVcObj=conTempSC.execute(query)
			if not rsGcVcObj.eof then
				vcHaveGcsTest=1
				exit for
			end if
			conTempSC.Close
			Set conTempSC=nothing
		end if
	next	
	%>
<%IF Instr(Ucase(Request.ServerVariables("SCRIPT_NAME")),"ONEPAGECHECKOUT")=0 THEN%>
	<div id="pcShowCartSmall">
		<a href="viewcart.asp"><img src="images/pc11-icon-cart.png" width="15" height="18" alt="Shopping Cart"></a>
		<%=dictLanguage.Item(Session("language")&"_addedtocart_5") & vcItems & dictLanguage.Item(Session("language")&"_smallcart_2")%>
		<a href="viewCart.asp"><%=dictLanguage.Item(Session("language")&"_smallcart_12") & scCurSign & money(vcTotal)%></a>
	</div>
<%ELSE
	opcOrderTotal=vcTotal
END IF%>
<%
'End Show shopping cart total
else
%>
	<a href="viewcart.asp"><img src="images/pc11-icon-cart.png" width="15" height="18" alt="Shopping Cart"></a>
    <%=dictLanguage.Item(Session("language")&"_smallcart_11")%>
<%
end if
%>