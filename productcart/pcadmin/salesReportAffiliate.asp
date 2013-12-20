<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="8*10*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"-->
<html>
<head>
<title>Affiliate Sales Report</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px;">
<div id="pcCPmain" style="background-image: none;">

	<% Dim conntemp, rs, query
	tempId=0
	
	' Choose the records to display
	err.clear
	Dim strTDateVar, strTDateVar2, DateVar, DateVar2
	strTDateVar=Request.queryString("FromDate")
	DateVar=strTDateVar
	if strTDateVar<>"" then
if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	if SQL_Format="1" then
	    DateVar=(DateVarArray(0)&"/"&DateVarArray(1)&"/"&DateVarArray(2))
	else
	    DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
	end if
else
    DateVarArray=split(strTDateVar,"/")
	if SQL_Format="1" then
	    DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
	else
	    DateVar=(DateVarArray(0)&"/"&DateVarArray(1)&"/"&DateVarArray(2))
	end if
end if
	end if
	strTDateVar2=Request.queryString("ToDate")
	DateVar2=strTDateVar2
	if strTDateVar2<>"" then
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	if SQL_Format = "1" then
	    DateVar2=(DateVarArray2(0)&"/"&DateVarArray2(1)&"/"&DateVarArray2(2))
	else
	    DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	end if
else
    DateVarArray2=split(strTDateVar2,"/")
	if SQL_Format = "1" then
	    DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	else
	    DateVar2=(DateVarArray2(0)&"/"&DateVarArray2(1)&"/"&DateVarArray2(2))
	end if
end if
	end if
	err.clear
	affVar=Request.queryString("idaffiliate1")
	if affVar="" then
		affVar=Request.queryString("idaffiliate2")
	end if
	
tmpDate=request("basedon")
tmpD=""
tmpD1=""
tmpD2=""
Select case tmpDate
Case "2": tmpD="orders.processDate"
tmpD1="processDate"
tmpD2="PROCESSED ON"
Case "3": tmpD="pcPackageInfo.pcPackageInfo_ShippedDate"
tmpD1="pcPackageInfo_ShippedDate"
tmpD2="SHIPPED ON"
Case Else: tmpD="orders.orderDate"
tmpD1="orderDate"
tmpD2="ORDER DATE"
End Select
	
	If DateVar<>"" then
		if scDB="Access" then
			query1=" AND " & tmpD & " >=#" & DateVar & "# "
		else
			query1=" AND " & tmpD & " >='" & DateVar & "' "
		end if
	else
		query1=""		
	End If

	If DateVar2<>"" then
		if scDB="Access" then
			query2=" AND " & tmpD & " <=#" & DateVar2 & "# "
		else
			query2=" AND " & tmpD & " <='" & DateVar2 & "' "
		end if
	else
		query2=""
	End If
	
	TempSpecial=""
	if tmpDate="3" then
	tmpStr1=""
	if query1<>"" then
		tmpStr1=replace(query1,tmpD,"orders.shipDate")
		tmpStr1=replace(tmpStr1," AND ","")
	end if
	tmpStr2=""
	if query2<>"" then
		tmpStr2=replace(query2,tmpD,"orders.shipDate")
		tmpStr2=replace(tmpStr2," AND ","")
	end if
	tmpD="orders.processDate"
	
	TempSpecial=" AND "
	if tmpStr1 & tmpStr2 <> "" then
		TempSpecial=TempSpecial & " ((" & tmpStr1
		if tmpStr2<>"" then
			if tmpStr1<>"" then
				TempSpecial=TempSpecial & " AND "
			end if
			TempSpecial=TempSpecial & tmpStr2 & ") OR "
		end if
	end if
	
	TempSpecial=TempSpecial & " (orders.idorder IN (SELECT DISTINCT idorder FROM pcPackageInfo"
	if query1<>"" or query2<>"" then
		TempSpecial=TempSpecial & " WHERE pcPackageInfo_ID>0 " & query1 & query2
	end if
	query1=""
	query2=""
	TempSpecial=TempSpecial & "))"
	if tmpStr1 & tmpStr2 <> "" then
		TempSpecial=TempSpecial & ")"
	end if
	end if

	If affVar="" OR affVar="0" then
		response.write "<div class=""pcCPmessage"">You must specify an affiliate to be able to generate a report. You can do this by either entering an ID or choosing one from the drop-down list. <a href=srcOrdByDate.asp>Back</a>.</div>"
		response.end
	End If
	
	call opendb()

	query="SELECT * FROM Orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & query1 & query2 & TempSpecial
				
	If affVar="ALL" then
		query=query&" AND idaffiliate>1 ORDER BY orders.idaffiliate, " & tmpD & ";"
	Else
		query=query&" AND idaffiliate="& affVar &" ORDER BY " & tmpD & ";"
	End if
	' Our Recordset Object
	Set rs=CreateObject("ADODB.Recordset")
	rs.CursorLocation=adUseClient
	rs.Open query, conntemp, 3, 3

	' If the returning recordset is not empty
	If rs.EOF Then
		set rs=nothing %>
		
		<table class="pcCPcontent">
			<tr> 
				<td><div class="pcCPmessage">Sorry, no records match your query</div></td>
			</tr>
		</table>

	<% Else %>

		<table border="0" cellspacing="0" cellpadding="4" width="100%" align="center" class="invoice">
			<tr>
				<td colspan="2">
					<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
						<td width="18%" height="71" valign="top"><img src="../pc/catalog/<%=scCompanyLogo%>"></td>
						<td width="39%" height="71" valign="top" align="center">
						<b><%=scCompanyName%></b><br>
						<%=scCompanyAddress%>, <%=scCompanyCity%>, <%=scCompanyState%>&nbsp;<%=scCompanyZip%><br>
						<%=scStoreURL%>
						</td>
						<td width="43%" height="71" valign="top" align="right">DATE: <%
						if scDateFrmt="DD/MM/YY" then
						    response.write day(Date()) & "/" & month(date()) & "/" & year(date)
						else
						    response.write month(Date()) & "/" & day(date()) & "/" & year(date)
						end if
						
						%></td>
					</table>
				</td>
			</tr>
			<tr> 
				<td width="50%" valign="top">
				
					<%
					If affVar<>"ALL" then
					query="SELECT idaffiliate,affiliateemail,affiliateName,affiliateAddress,affiliateAddress2,affiliatecity,affiliatestate,affiliatezip,affiliateCountryCode FROM affiliates WHERE idaffiliate="& affVar
					Set rsObjAff=CreateObject("ADODB.Recordset")
					set rsObjAff=conntemp.execute(query)
					intIdAffiliate=rsObjAff("idaffiliate")
					strAffiliateEmail=rsObjAff("affiliateemail")
					strAffiliateName=rsObjAff("affiliateName")
					strAffiliateAddress=rsObjAff("affiliateAddress")
					strAffiliateAddress2=rsObjAff("affiliateAddress2")
					strAffiliateCity=rsObjAff("affiliatecity")
					strAffiliateState=rsObjAff("affiliatestate")
					strAffiliateZip=rsObjAff("affiliatezip")
					strAffiliateCountryCode=rsObjAff("affiliateCountryCode")
					set rsObjAff=nothing
					%>
				 
							Affiliate ID: #<%=affVar%><br>
							AFFILIATE NAME:	<%=strAffiliateName%><br>
							<%=strAffiliateAddress%><br>
							<% if strAffiliateAddress2<>"" then
								response.write strAffiliateAddress2&"<BR>"
							end if %>  
							<%=strAffiliateCity%>, <%=strAffiliateState%>&nbsp;<%=strAffiliateZip%>
							<%if strAffiliateCountryCode <> scShipFromPostalCountry then
								response.write "<BR>" & strAffiliateCountryCode
							end if %>
							<br>E-mail: <%=strAffiliateEmail%>

					<% end if %>
			</td>
			<td valign="top">
					<b>
					<%if (strTDateVar<>"") or (strTDateVar2<>"") then%>Total Sales 
					recorded<%if strTDateVar<>"" then%>&nbsp;from:&nbsp;<%=strTDateVar%><%end if%><%if strTDateVar2<>"" then%>&nbsp;to:&nbsp;<%=strTDateVar2%><%end if%><br><%end if%></b>
					<% Response.Write "Total Records Found : " & rs.RecordCount & "<br><br>"%>
			</td>
		</tr>
	</table>
	
	<br>
	
	<table border="0" cellspacing="0" cellpadding="4" width="100%" align="center" class="invoice">
		<tr> 
			<td valign="top" align="left" width="12%" nowrap><b><%=tmpD2%></b></td>
			<td valign="top" align="left" width="8%" nowrap><b>ORDER #</b></td>
			<td valign="top" align="left" width="60%"><b>ORDER DETAILS</b></td>
			<td valign="top" align="left" nowrap width="10%"><b>TOTAL SALES</b></td>
			<td valign="top" align="left" nowrap width="10%"><b>ORDER TOTAL</b></td>
			<td valign="top" align="left" nowrap width="10%"><b>SHIPPING</b></td>
			<td valign="top" align="left" nowrap width="10%"><b>TAX</b></td>
			<td valign="top" align="left" nowrap width="10%"><b>COMMISSION</b></td>
		 </tr>
		<% 
		gTotalsales=0
		gTotalOrder=0
		gTotalShip=0
		gTotalTax=0
		gTotalcomm=0
		do until rs.EOF
			gSubOrder=0
			gSubTax=0
			gSubShip=0
			gSubCom=0
			intIdaffiliate=rs("idaffiliate")
			intIdOrder=rs("idOrder")
			intIdCustomer=rs("idcustomer")
			'dblTotal=rs("total")
			'dblTaxAmount=rs("taxAmount")
			if tmpDate<>"3" then
				dtOrderDate=rs(tmpD1)
				if scDateFrmt="DD/MM/YY" then
					dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
				end if
			else
				call opendb()
				query="SELECT pcPackageInfo_ShippedDate FROM pcPackageInfo WHERE idorder=" & intIdOrder
				set rsStr=connTemp.execute(query)
				dtOrderDate=""
				if not rsStr.eof then
					do while not rsStr.eof
						tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
						if scDateFrmt="DD/MM/YY" then
							tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
						end if
						if  dtOrderDate<>"" then
							dtOrderDate=dtOrderDate & "<br>"
						end if
						dtOrderDate=dtOrderDate & tmp_processDate
						rsStr.MoveNext
					loop
				else
					query="SELECT shipDate FROM orders WHERE idorder=" & intIdOrder
					set rsStr=connTemp.execute(query)
					if not rsStr.eof then
						dtOrderDate=rsStr("shipDate")
						if scDateFrmt="DD/MM/YY" then
							dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
						end if
					end if
				end if
				set rsStr=nothing
			end if
			dblAffiliatePay=rs("affiliatePay")
			porderdetails=rs("details")
			tmp1=split(porderdetails,"Amount: ||")
			porderdetails=""
			For k=lbound(tmp1) to ubound(tmp1)
				if trim(tmp1(k))<>"" then
				tmp2=trim(tmp1(k))
				tmp2=mid(tmp2,instr(tmp2," ")+1,len(tmp2))
				if k<>ubound(tmp1) then
				porderdetails=porderdetails & tmp2 & "<br>"
				else
				porderdetails=porderdetails & tmp2
				end if
				end if
			Next

'Calculate "NET" Order Amount
ptotal=rs("total")
gSubOrder=rs("total")
ptaxAmount=rs("taxAmount")
ptaxDetails=rs("taxDetails")
pord_VAT=rs("ord_VAT")
gSubTax=ptaxAmount+pord_VAT
pshipmentDetails=rs("shipmentDetails")
Postage=0
serviceHandlingFee=0
shipping=split(pshipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
		else
			Postage=cdbl(trim(shipping(2)))
			if ubound(shipping)=>3 then
				serviceHandlingFee=cdbl(trim(shipping(3)))
				if NOT isNumeric(serviceHandlingFee) then
					serviceHandlingFee=0
				end if
			else
				serviceHandlingFee=0
			end if
		end if
	end if
gSubShip=Postage
ppaymentDetails=trim(rs("paymentDetails"))
payment = split(ppaymentDetails,"||")
PayCharge=0
If ubound(payment)>=1 then					
If payment(1)="" then
	PayCharge=0
else
	PayCharge=payment(1)
end If
End if

PrdSales=ptotal
PrdSales=PrdSales-postage
PrdSales=PrdSales-serviceHandlingFee
PrdSales=PrdSales-PayCharge

gSubOrder=gSubOrder-postage


pdiscountDetails=rs("discountDetails")
pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
	pcv_CatDiscounts="0"
end if

if (instr(pdiscountDetails,"- ||")>0) or (pcv_CatDiscounts>"0")  then
	DiscountDetailsArry=split(pdiscountDetails,",")
	intArryCnt=ubound(DiscountDetailsArry)
		
	dim discounts, discountType 
						
	discount=0
	for k=0 to intArryCnt
		pTempDiscountDetails=DiscountDetailsArry(k)
		if instr(pTempDiscountDetails,"- ||") then
			discounts = split(pTempDiscountDetails,"- ||")
			tdiscount = discounts(1)
		else
			tdiscount=0
		end if
		discount=discount+tdiscount
	Next
	PrdSales=PrdSales+discount+pcv_CatDiscounts
end if

if pord_VAT>0 then
	PrdSales=PrdSales-pord_VAT
else
	if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
		PrdSales=PrdSales-ptaxAmount
	else
		if cdbl(ptaxAmount)=0 then
			taxArray=split(ptaxDetails,",")
			for i=0 to (ubound(taxArray)-1)
				taxDesc=split(taxArray(i),"|")
				PrdSales=PrdSales-taxDesc(1)
				gSubTax=gSubTax+taxDesc(1)
			next 
		else
			PrdSales=PrdSales-ptaxAmount
		end if
	end if
end if

gSubOrder=gSubOrder-gSubTax

			
			query="SELECT idaffiliate,affiliateemail,affiliateName,affiliateAddress,affiliateAddress2,affiliatecity,affiliatestate,affiliatezip,affiliateCountryCode FROM affiliates WHERE idaffiliate="& intIdaffiliate
			Set rsObjAff=CreateObject("ADODB.Recordset")
			rsObjAff.Open query, scDSN , 3, 3
			
			query="SELECT unitPrice,quantity FROM ProductsOrdered WHERE idorder="& intIdOrder
			Set rstemp=CreateObject("ADODB.Recordset")
			rstemp.CursorLocation=adUseClient
			rstemp.Open query, scDSN , 3, 3
			bOrderTotal=0
			do until rstemp.eof
				unitTotal=rstemp("unitPrice")
				quantity=rstemp("quantity")
				query="SELECT name,lastname FROM customers WHERE idcustomer="& intIdcustomer
				Set rsCust=CreateObject("ADODB.Recordset")
				rsCust.CursorLocation=adUseClient
				rsCust.Open query, scDSN , 3, 3
				CustName=rsCust("name")& " "&rsCust("lastname")
				rsCust.Close
				set rsCust=nothing
		
				bOrderTotal=0 + (unitTotal * quantity)
				rstemp.moveNext
			loop
			rstemp.Close
			set rstemp=nothing
			
			gTotalsales=gTotalsales + PrdSales
			gTotalOrder=gTotalOrder+gSubOrder
			gTotalShip=gTotalShip+gSubShip
			gTotalTax=gTotalTax+gSubTax
			
			If affVar="ALL" then 
				intIdAffiliate=rsObjAff("idaffiliate")
				strAffiliateEmail=rsObjAff("affiliateemail")
				strAffiliateName=rsObjAff("affiliateName")
				strAffiliateAddress=rsObjAff("affiliateAddress")
				strAffiliateAddress2=rsObjAff("affiliateAddress2")
				strAffiliateCity=rsObjAff("affiliatecity")
				strAffiliateState=rsObjAff("affiliatestate")
				strAffiliateZip=rsObjAff("affiliatezip")
				strAffiliateCountryCode=rsObjAff("affiliateCountryCode")
				rsObjAff.Close
				set rsObjAff=nothing
				if tempId<>intIdAffiliate then
					tempId=intIdAffiliate %>
					<tr> 
						<td colspan="8">
							<b><img src="images/pc_individual.gif" width="14" height="15"> 
							Affiliate ID: <%=intIdAffiliate%>&nbsp;<%=strAffiliateName%> 
						- </b><%=strAffiliateAddress%> 
							<% if len(strAffiliateAddress2)>0 then%>
								<%response.write ", " & strAffiliateAddress2%>
							<% end If %>, 
							<%=strAffiliateCity%>, <%=strAffiliateState%>&nbsp;<%=strAffiliateZip%>&nbsp;<%=strAffiliateCountryCode%>
						</td>
					</tr>
				<% end if 
			end if %>
			<tr> 
				<td height="21" width="12%" nowrap valign="top"><%=dtOrderDate%></td>
				<td height="21" width="8%" nowrap valign="top"><%=(scpre+int(intIdOrder))%></td>
				<td height="21" width="60%" valign="top"><%=porderdetails%></td>
				<td height="21" width="10%" nowrap align="right" valign="top"><%=scCurSign&money(PrdSales)%></td>
				<td height="21" width="10%" nowrap align="right" valign="top"><%=scCurSign&money(gSubOrder)%></td>
				<td height="21" width="10%" nowrap align="right" valign="top"><%=scCurSign&money(gSubShip)%></td>
				<td height="21" width="10%" nowrap align="right" valign="top"><%=scCurSign&money(gSubTax)%></td>
				<td height="21" width="10%"  nowrap align="right" valign="top"><%=scCurSign&money(dblAffiliatePay)%></td>
			</tr>
			<% gTotalcomm=gTotalcomm + dblAffiliatePay %>
			<% rs.MoveNext
		loop
		set rs=nothing
	End If %>
	</table>
	
	<br>

	<table width="100%" border="0" align="center" cellpadding="4" cellspacing="0" class="invoice">
	<tr bgcolor="#e1e1e1">
		<td colspan="3">&nbsp;</td>
		<td nowrap> <div align="right"><b>Total Sales</b></div></td>
		<td nowrap> <div align="right"><b>Total Orders</b></div></td>
		<td nowrap> <div align="right"><b>Total Shipping</b></div></td>
		<td nowrap> <div align="right"><b>Total Taxes</b></div></td>
		<td nowrap> <div align="right"><b>Total Commissions</b></div></td>
	</tr>
	<tr> 
		<td height="21" colspan="3">&nbsp;</td>
		<td height="21" width="10%" nowrap> <div align="right"><b><%=scCurSign&money(gTotalsales)%></b></div></td>
		<td height="21" width="10%" nowrap> <div align="right"><b><%=scCurSign&money(gTotalOrder)%></b></div></td>
		<td height="21" width="10%" nowrap> <div align="right"><b><%=scCurSign&money(gTotalShip)%></b></div></td>
		<td height="21" width="10%" nowrap> <div align="right"><b><%=scCurSign&money(gTotalTax)%></b></div></td>
		<td height="21" width="10%"  nowrap> <div align="right"><b><%=scCurSign&money(gTotalcomm)%></b></div></td>
	</tr>
</table>
<%	' Done. Now release Objects
	call closedb()
%>
</div>
</body>
</html>