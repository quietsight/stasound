<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Sales Report" %>
<% section="" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->  
<!--#include file="../includes/SQLFormat.txt"-->
<% 	response.Buffer=true
	
if(Request.Form("ReturnAS")="Content") then
	Response.ContentType="application/msexcel"
end if
Response.Expires=0
	
on error resume next 
dim query, conntemp, rstemp
call opendb()
' Choose the records to display
err.clear

Function GetProductCFs(idOrder,tmpOrdDetails)
Dim tmp1,queryQ,rsQ,tmpArr,intCount,i,j
	tmp1=split(tmpOrdDetails,vbcrlf)
	queryQ="SELECT Products.sku,Products.Description,ProductsOrdered.xfdetails FROM Products INNER JOIN ProductsOrdered ON Products.idProduct=ProductsOrdered.idProduct WHERE ProductsOrdered.idOrder=" & idOrder & " AND ProductsOrdered.xfdetails<>'';"
	set rsQ=connTemp.execute(queryQ)
	if not rsQ.eof then
		tmpArr=rsQ.getRows()
		intCount=ubound(tmpArr,2)
		For i=0 to intCount
			For j=lbound(tmp1) to ubound(tmp1)
				If (Instr(tmp1(j),tmpArr(0,i))>0) AND (Instr(tmp1(j),tmpArr(1,i))>0) then
					tmp1(j)=tmp1(j) & vbcrlf & replace(tmpArr(2,i),"|",vbcrlf)
					Exit For
				End if
			Next
		Next
	end if
	set rsQ=nothing
	GetProductCFs=join(tmp1,vbcrlf)
End Function

Dim strTDateVar, strTDateVar2, DateVar, DateVar2
strTDateVar=Request.Form("FromDate")
DateVar=strTDateVar
if (strTDateVar<>"") and (isDate(strTDateVar)) then
	if scDateFrmt="DD/MM/YY" then
		DateVarArray=split(strTDateVar,"/")
		DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
	end if
end if
strTDateVar2=Request.Form("ToDate")
DateVar2=strTDateVar2
if (strTDateVar2<>"") and (isDate(strTDateVar2)) then
	if scDateFrmt="DD/MM/YY" then
		DateVarArray2=split(strTDateVar2,"/")
		DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
		if err.number<>0 then
			DateVar=Request.Form("FromDate")
			DateVar2=Request.Form("ToDate")
		end if
	end if
end if
err.clear

tmpDate=request("basedon")
tmpD=""
tmpD1=""
tmpD2=""
Select case tmpDate
Case "2": tmpD="orders.processDate"
tmpD1="processDate"
tmpD2="Processed On"
Case "3": tmpD="pcPackageInfo.pcPackageInfo_ShippedDate"
tmpD1="pcPackageInfo_ShippedDate"
tmpD2="Shipped On"
Case Else: tmpD="orders.orderDate"
tmpD1="processDate"
tmpD2="Processed On"
End Select
	
If DateVar<>"" then
	if SQL_Format="1" then
		DateVar=Day(DateVar)&"/"&Month(DateVar)&"/"&Year(DateVar)
	end if

	if scDB="Access" then
		query1=" AND " & tmpD & " >=#" & DateVar & "# "
	else
		query1=" AND " & tmpD & " >='" & DateVar & "' "
	end if
else
	query1=""		
End If

If DateVar2<>"" then
	if SQL_Format="1" then
		DateVar2=Day(DateVar2)&"/"&Month(DateVar2)&"/"&Year(DateVar2)
	end if
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

strSQL="SELECT orders.idOrder,orderDate,idCustomer,details,total,address,zip,stateCode,state,city,countryCode,comments,taxAmount,shipmentDetails,paymentDetails,discountDetails,randomNumber,shippingAddress,shippingStateCode,shippingState,shippingCity,shippingCountryCode,shippingZip,orderStatus,pcOrd_shippingPhone,idAffiliate,processDate,shipDate,shipVia,trackingNum,affiliatePay,returnDate,returnReason,iRewardPoints,ShippingFullName,iRewardValue,iRewardRefId,iRewardPointsRef,iRewardPointsCustAccrued,IDRefer,address2,shippingCompany,shippingAddress2,taxDetails,adminComments,rmaCredit,DPs,gwAuthCode,gwTransId,paymentCode,SRF,ordShiptype,ordPackageNum,ord_DeliveryDate,ord_OrderName,ord_VAT,pcOrd_CatDiscounts FROM Orders WHERE ((orders.orderStatus>1 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & query1 & query2 & TempSpecial & " ORDER BY " & tmpD & ";"

set rstemp=Server.CreateObject("ADODB.Recordset") 
set rstemp=conntemp.execute(strSQL)
if not rstemp.eof then
	DataEmpty=0
else
	DataEmpty=1
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
	<head>
		<title>Sales Report from <%=DateVar%> to <%=DateVar2%></title>
        <style>
		h1 {
			font-family: Arial, Helvetica, sans-serif;
			font-size: 16px;
			font-weight: bold;
		}
		
		table.salesExport {
			padding: 0;
			margin: 0;
		}
		
		table.salesExport td {
			font-family: Arial, Helvetica, sans-serif;
			font-size: 11px;
			padding: 3px;
			border-right: 1px solid #CCC;
			border-bottom: 1px solid #CCC;
		}
		
		table.salesExport th {
			font-family: Arial, Helvetica, sans-serif;
			font-size: 12px;
			padding: 3px;
			font-weight: bold;
			text-align: left;
			background-color: #f5f5f5;
			border-right: 1px solid #CCC;
			border-bottom: 1px solid #CCC;
		}
		</style>
	</head>
	<body>
	<% dim strReturnAs
	strReturnAs=request.Form("ReturnAS")
	select case strReturnAS
		case "CSV"
			CreateCSVFile()
		case "HTML"
			GenHTML()
		case "XLS"
			CreateXlsFile()
	end select
   
	Response.Flush
	%>
	</body>
</html>

<% Function GenFileName()
	dim fname
	
	fname="File"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime))
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

Function GenHTML()
	Response.Write("<h1>Sales Report from " & DateVar & " to " & DateVar2 & "</h1>")
	Response.Write("<TABLE class=salesExport>")
	Response.Write("<TR valign='top'>")
	Response.Write("	<Th nowrap align='right'>Order #</Th>")
	Response.Write("	<Th nowrap>Order Name</Th>")
	Response.Write("	<Th nowrap>Order Date</Th>")
	Response.Write("	<Th nowrap>Order Status</Th>")
	Response.Write("	<Th nowrap style='width: 300px;'>Order Details</Th>")
	Response.Write("	<Th nowrap>Processed Date</Th>")
	Response.Write("	<Th nowrap>Customer</Th>")
	Response.Write("	<Th nowrap>Shipment Address</Th>")
	Response.Write("	<Th nowrap>Shipment Details</Th>")
	Response.Write("	<Th nowrap>Payment Details</Th>")
	Response.Write("	<Th nowrap>Discount Details</Th>")
	Response.Write("	<Th nowrap>Tax Details</Th>")
	Response.Write("	<Th nowrap align='right'>VAT</Th>")
	Response.Write("	<Th nowrap align='right'>Credit</Th>")
	if RewardsActive = 1 then
	Response.Write("	<Th nowrap align='right'>" & RewardsLabel & "</Th>")
	end if
	Response.Write("	<Th nowrap>Affiliate</Th>")
	Response.Write("	<Th nowrap>Referrer</Th>")
	Response.Write("	<Th nowrap align='right'>Order Total</Th>")	
	Response.Write("	<Th nowrap>Customer Comments</Th>")
	Response.Write("	<Th nowrap>Admin Comments</Th>")
	Response.Write("</TR>")
	if(DataEmpty=1) then
		Response.Write("Database Empty")
	else
		do until rstemp.eof
			pcv_idOrder=rstemp(0)
			pcv_orderDate=rstemp(1)
			pcv_idCustomer=rstemp(2)
			pcv_details=rstemp(3)
			pcv_total=rstemp(4)
			if pcv_total<>"" then
			else
			pcv_total="0"
			end if
			pcv_address=rstemp(5)
			pcv_zip=rstemp(6)
			pcv_stateCode=rstemp(7)
			pcv_state=rstemp(8)
			pcv_city=rstemp(9)
			pcv_countryCode=rstemp(10)
			pcv_comments=rstemp(11)
			pcv_taxAmount=rstemp(12)
			if pcv_taxAmount<>"" then
			else
			pcv_taxAmount="0"
			end if
			pcv_shipmentDetails=rstemp(13)
			pcv_paymentDetails=rstemp(14)
			pcv_discountDetails=rstemp(15)
			pcv_randomNumber=rstemp(16)
			pcv_shippingAddress=rstemp(17)
			pcv_shippingStateCode=rstemp(18)
			pcv_shippingState=rstemp(19)
			pcv_shippingCity=rstemp(20)
			pcv_shippingCountryCode=rstemp(21)
			pcv_shippingZip=rstemp(22)
			pcv_orderStatus=rstemp(23)
			pcv_shippingPhone=rstemp(24)
			pcv_idAffiliate=rstemp(25)
			pcv_processDate=rstemp(26)
			
			query="SELECT pcPackageInfo_ShippedDate FROM pcPackageInfo WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			pcv_shipDate=""
			if not rsStr.eof then
			do while not rsStr.eof
				tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
				if scDateFrmt="DD/MM/YY" then
					tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
				end if
				if  pcv_shipDate<>"" then
					pcv_shipDate=pcv_shipDate & "<br>"
				end if
				pcv_shipDate=pcv_shipDate & tmp_processDate
				rsStr.MoveNext
			loop
			else
			query="SELECT shipDate FROM orders WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			if not rsStr.eof then
				pcv_shipDate=rsStr("shipDate")
				if scDateFrmt="DD/MM/YY" then
					pcv_shipDate=(day(pcv_shipDate)&"/"&month(pcv_shipDate)&"/"&year(pcv_shipDate))
				end if
			end if
			end if
			set rsStr=nothing
			
			pcv_shipVia=rstemp(28)
			pcv_trackingNum=rstemp(29)
			pcv_affiliatePay=rstemp(30)
			pcv_returnDate=rstemp(31)
			pcv_returnReason=rstemp(32)
			pcv_iRewardPoints=rstemp(33)
			if pcv_iRewardPoints<>"" then
			else
			pcv_iRewardPoints="0"
			end if
			pcv_ShippingFullName=rstemp(34)
			pcv_iRewardValue=rstemp(35)
			pcv_iRewardRefId=rstemp(36)
			pcv_iRewardPointsRef=rstemp(37)
			pcv_iRewardPointsCustAccrued=rstemp(38)
			pcv_IDRefer=rstemp(39)
			if pcv_IDRefer<>"" then
			else
			pcv_IDRefer="0"
			end if
			pcv_address2=rstemp(40)
			pcv_shippingCompany=rstemp(41)
			pcv_shippingAddress2=rstemp(42)
			pcv_taxDetails=rstemp(43)
			pcv_adminComments=rstemp(44)
			pcv_rmaCredit=rstemp(45)
			if pcv_rmaCredit<>"" then
			else
			pcv_rmaCredit="0"
			end if
			pcv_DPs=rstemp(46)
			pcv_gwAuthCode=rstemp(47)
			pcv_gwTransId=rstemp(48)
			pcv_paymentCode=rstemp(49)
			pcv_SRF=rstemp(50)
			pcv_ordShiptype=rstemp(51)
			if pcv_ordShiptype<>"" then
			else
			pcv_ordShiptype="0"
			end if
			pcv_ordPackageNum=rstemp(52)
			if pcv_ordPackageNum<>"" then
			else
			pcv_ordPackageNum="0"
			end if
			pcv_ord_DeliveryDate=rstemp(53)
			pcv_ord_OrderName=rstemp(54)
			pcv_ord_VAT=rstemp(55)
			if pcv_ord_VAT<>"" then
			else
			pcv_ord_VAT="0"
			end if
			
			pcv_pcOrd_CatDiscounts=rstemp(56)
			if pcv_pcOrd_CatDiscounts<>"" then
			else
			pcv_pcOrd_CatDiscounts="0"
			end if
					 
			Response.Write("<TR valign='top'>")
	
			'***** Column 1 *****
			Response.Write("<TD align='right' nowrap>")
			Response.Write(int(pcv_idorder+scPre))
			Response.Write("&nbsp;</TD>")
			
			'***** Column 2 *****
			Response.Write("<TD>")
			Response.Write(pcv_ord_OrderName)
			Response.Write("&nbsp;</TD>")
					
			'***** Column 3 *****
			Response.Write("<TD nowrap>")
			dtOrderDate=pcv_orderDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			Response.Write(dtOrderDate)
			Response.Write("&nbsp;</TD>")
	
			'***** Column 4 *****
			if pcv_orderStatus="2" then
				strOStatus="Pending"
			end if
			if pcv_orderStatus="3" then
				strOStatus="Processed"
			end if
			if pcv_orderStatus="4" then
				strOStatus="Shipped"
			end if
			if pcv_orderStatus="7" then
				strOStatus="Partially Shipped"
			end if
			if pcv_orderStatus="8" then
				strOStatus="Shipping"
			end if
			if pcv_orderStatus="10" then
				strOStatus="Delivered"
			end if
			if pcv_orderStatus="12" then
				strOStatus="Archived"
			end if
			Response.Write("<TD nowrap>")
			Response.Write(strOStatus)
			Response.Write("&nbsp;</TD>")
			
			'***** Column 5 *****
			Response.Write("<TD width='400'>")
			Response.Write(replace(replace(GetProductCFs(pcv_idOrder,pcv_details)," ||",""&scCurSign&""),vbcrlf,"<br>"))
			Response.Write("&nbsp;</TD>")
			
			'***** Column 6 *****
			Response.Write("<TD nowrap>")
			dtOrderDate=pcv_processDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			Response.Write(dtOrderDate)
			Response.Write("&nbsp;</TD>")
			
			'***** Column 7 *****
			query="select name,lastname,customerCompany,phone FROM customers where idcustomer=" & pcv_idCustomer
			set rs=connTemp.execute(query)
			
			pcv_CustName=rs("name") & " " & rs("lastname") & " (#" & pcv_idcustomer & ")"
			pcv_CustCompany=rs("customerCompany")
			pcv_CustPhone=rs("phone")
			Response.Write("<TD nowrap>")
			Response.Write(pcv_CustName & "<br>")
			if trim(pcv_CustCompany)<>"" then
				Response.Write("Company: " & pcv_CustCompany & "<br>")
			end if
			Response.Write(pcv_address & "<br>")
			if trim(pcv_address2)<>"" then
				Response.Write("Address 2: " & pcv_address2 & "<br>")
			end if
			Response.write(pcv_city &", " & pcv_statecode & pcv_state & " " & pcv_zip & ", " & pcv_countryCode & "<br>")
			if pcv_CustPhone<>"" then
				Response.Write("Phone: " & pcv_CustPhone)
			end if
			Response.Write("&nbsp;</TD>")
			
			'***** Column 8 *****
			Response.Write("<TD nowrap>")
			if pcv_ShippingFullName<>"" then
				Response.Write("Shipping Name: " & pcv_ShippingFullName & "<br>")
			end if
			if pcv_shippingCompany<>"" then
				Response.Write("Company: " & pcv_shippingCompany & "<br>")
			end if
			Response.Write(pcv_shippingAddress & "<br>")
			if pcv_shippingAddress2<>"" then
				Response.Write("Address 2: " & pcv_shippingAddress2 & "<br>")
			end if
			Response.write(pcv_shippingCity &", " & pcv_shippingStatecode & pcv_shippingState & " " & pcv_shippingZip & ", " & pcv_shippingCountryCode & "<br>")
			if trim(pcv_shippingPhone)<>"" then
				Response.Write("Phone: " & pcv_shippingPhone)
			end if
			Response.Write("&nbsp;</TD>")
	
			'***** Column 9 *****
			ShipDetails=""
			if pcv_SRF="1" then
			ShipDetails="Shipping charges to be determined."
			else
			if instr(pcv_shipmentDetails,",")=0 then
			ShipDetails=pcv_shipmentDetails
			end if			
			end if
			
			ShipFees=""
			HandlingFees=""
			if ShipDetails="" then
				pcv_Ship=split(pcv_shipmentDetails,",")
				ShipDetails=pcv_Ship(1)
				ShipFees=pcv_Ship(2)
				HandlingFees=pcv_Ship(3)
			end if
			
			Response.Write("<TD nowrap>")
			Response.Write("Shipping Method: " & ShipDetails & "<br>")
			if pcv_ordShiptype="0" then
			Response.Write("Shipping Type: Residential <br>")
			end if
			if pcv_ordShiptype="1" then
			Response.Write("Shipping Type: Commercial <br>")
			end if
			if ShipFees>"0" then
			Response.Write("Fees: " & scCurSign & money(ShipFees) & "<br>")
			end if
			if HandlingFees>"0" then
			Response.Write("Handling Fees: " & scCurSign & money(HandlingFees) & "<br>")
			end if
			if pcv_ShipVia<>"" then
			Response.Write("Shipped Via: " & pcv_ShipVia & "<br>")
			end if
			if pcv_trackingNum<>"" then
			Response.Write("Number of packages: " & pcv_trackingNum & "<br>")
			end if
			if pcv_shipDate<>"" then
				dtOrderDate=pcv_shipDate
				Response.Write("Date Shipped: " & dtOrderDate & "<br>")
			end if
			if pcv_ord_DeliveryDate<>"" then
				dtOrderDate=pcv_ord_DeliveryDate
				if scDateFrmt="DD/MM/YY" then
					dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
				end if
				Response.Write("Delivery Date: " & dtOrderDate)
			end if
	
			Response.Write("&nbsp;</TD>")
	
			'***** Column 10 *****
			Response.Write("<TD nowrap>")
			if instr(pcv_paymentDetails,"||")>0 then
				pcv_PayArray=split(pcv_paymentDetails,"||")
				Response.Write("Payment Method: " & trim(pcv_PayArray(0)) & "<br>")
				if trim(pcv_PayArray(1))<>"" then
				if IsNumeric(trim(pcv_PayArray(1))) then
					PayFees=cdbl(trim(pcv_PayArray(1)))
					if PayFees>0 then
						Response.Write("Fees: " & scCurSign & money(PayFees) & "<br>")
					end if
				end if
				end if
			else
				Response.Write("Payment Details: " & pcv_paymentDetails & "<br>")
			end if
			if pcv_paymentCode<>"" then
			Response.Write("Payment Gateway: " & pcv_paymentCode & "<br>")
			end if
			if pcv_gwTransId<>"" then
			Response.Write("Transaction ID: " & pcv_gwTransId & "<br>")
			end if
			if pcv_gwAuthCode<>"" then
			Response.Write("Authorization Code: " & pcv_gwAuthCode)
			end if
			Response.Write("&nbsp;</TD>")
			
			'***** Column 11 *****
			Response.Write("<TD nowrap>")
			if instr(pcv_discountDetails,"- ||")>0 then
				tmpDArr=split(pcv_discountDetails,"- ||")
				For m=lbound(tmpDArr)+1 to ubound(tmpDArr)
					if instr(tmpDArr(m),",")>0 then
						tmpDArr(m)=replace(tmpDArr(m),",","mmmmm",1,1)
					end if
				Next
				pcv_discountDetails=Join(tmpDArr,"- ||")
				pcv_DisArray1=split(pcv_discountDetails,"mmmmm")
			For m=lbound(pcv_DisArray1) to ubound(pcv_DisArray1)
			if pcv_DisArray1(m)<>"" then
				pcv_DisArray=split(pcv_DisArray1(m),"- ||")
				DisAmount=cdbl(trim(pcv_DisArray(1)))
				Response.Write("Discount Name: " & trim(pcv_DisArray(0)) & "<br>")
				if DisAmount<>0 then
					Response.Write("Amount: -" & scCurSign & money(DisAmount) & "<br>")
				end if
			end if
			Next
			end if
	
			if pcv_pcOrd_CatDiscounts<>"" then
				if cdbl(pcv_pcOrd_CatDiscounts)<>0 then
					Response.Write("Discount by Categories: -" & scCurSign & money(pcv_pcOrd_CatDiscounts))
				end if
			end if
			Response.Write("&nbsp;</TD>")
			
			'***** Column 12 *****
			Response.Write("<TD nowrap>")
			Response.Write("Tax Amount: " & scCurSign & money(pcv_taxAmount) & "<br>")
			IF cdbl(pcv_taxAmount)>0 then
				if instr(pcv_taxDetails,",")>0 then
				TaxArray=split(pcv_taxDetails,",")
				For m=0 to (ubound(TaxArray)-1)
					tmpTax=split(TaxArray(m),"|")
					Response.Write(ucase(tmpTax(0)) & " - Amount: " & scCurSign & money(tmpTax(1)) & "<br>")
				Next
				end if
			END IF
			Response.Write("&nbsp;</TD>")
			
			'***** Column 13 *****
			Response.Write("<TD align='right' nowrap>")
			Response.Write(scCurSign & money(pcv_ord_VAT))
			Response.Write("&nbsp;</TD>")
			
			'***** Column 14 *****
			Response.Write("<TD align='right' nowrap>")
			if pcv_rmaCredit>"0" then
			Response.Write("-" & scCurSign & money(pcv_rmaCredit))
			end if
			Response.Write("&nbsp;</TD>")
			
			'***** Column 15 *****
			if RewardsActive = 1 then
			Response.Write("<TD align='right' nowrap>")
			if pcv_iRewardPoints>"0" then
			Response.Write(scCurSign & money(pcv_iRewardPoints))
			end if
			Response.Write("&nbsp;</TD>")
			end if
	
			'***** Column 16 *****
			Response.Write("<TD nowrap>")
			if pcv_idAffiliate>"1" then
				query="select affiliateName,commission FROM affiliates where idAffiliate=" & pcv_idAffiliate
				set rs=connTemp.execute(query)
				if not rs.eof then
					Response.Write(rs("affiliateName") & " (#" & pcv_idAffiliate & ")" & "<br>")
					if pcv_affiliatePay>"0" then
						Response.Write("Amount: " & scCurSign & money(pcvTotal*(rs("commission")/100)))
					end if
				end if
			end if
			Response.Write("&nbsp;</TD>")
	
			'***** Column 17 *****
			Response.Write("<TD nowrap>")
			if pcv_idRefer>"0" then
				query="select Name FROM Referrer where IdRefer=" & pcv_idRefer
				set rs=connTemp.execute(query)
				Response.Write(rs("Name") & " (#" & pcv_idRefer & ")" )
			end if
			Response.Write("&nbsp;</TD>")
			
			'***** Column 18 *****
			Response.Write("<TD align='right' nowrap>")
			Response.Write(scCurSign & money(pcv_total))
			Response.Write("&nbsp;</TD>")
			
			'***** Column 19 *****
			Response.Write("<TD width='250'>")
			Response.Write(pcv_comments)
			Response.Write("&nbsp;</TD>")
			
			'***** Column 20 *****
			Response.Write("<TD width='250'>")
			Response.Write(pcv_adminComments)
			Response.Write("&nbsp;</TD>")
			
			Response.Write("</TR>")
			rstemp.moveNext
		loop
		set rstemp=nothing
		Response.Write("</TABLE>")
	End if
End Function

Function CreateCSVFile()
	
	set StringBuilderObj = new StringBuilder	
	
	strFile=GenFileName()   
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
	if (DataEmpty=0) then
		StringBuilderObj.append chr(34) & "Order #" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Order Name" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Order Date" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Order Status" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Order Details" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Processed Date" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Customer" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Shipment Address" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Shipment Details" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Payment Details" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Discount Details" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Tax Details" & chr(34) & ","
		StringBuilderObj.append chr(34) & "VAT" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Credit" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Reward Points" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Affiliate" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Referrer" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Order Total" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Customer Comments" & chr(34) & ","
		StringBuilderObj.append chr(34) & "Admin Comments" & chr(34) & ","
		
		a.WriteLine(StringBuilderObj.toString())
		
		set StringBuilderObj = nothing
		
		do until rstemp.eof
			pcv_idOrder=rstemp(0)
			pcv_orderDate=rstemp(1)
			pcv_idCustomer=rstemp(2)
			pcv_details=rstemp(3)
			pcv_total=rstemp(4)
			if pcv_total<>"" then
			else
			pcv_total="0"
			end if
			pcv_address=rstemp(5)
			pcv_zip=rstemp(6)
			pcv_stateCode=rstemp(7)
			pcv_state=rstemp(8)
			pcv_city=rstemp(9)
			pcv_countryCode=rstemp(10)
			pcv_comments=rstemp(11)
			pcv_taxAmount=rstemp(12)
			if pcv_taxAmount<>"" then
			else
			pcv_taxAmount="0"
			end if
			pcv_shipmentDetails=rstemp(13)
			pcv_paymentDetails=rstemp(14)
			pcv_discountDetails=rstemp(15)
			pcv_randomNumber=rstemp(16)
			pcv_shippingAddress=rstemp(17)
			pcv_shippingStateCode=rstemp(18)
			pcv_shippingState=rstemp(19)
			pcv_shippingCity=rstemp(20)
			pcv_shippingCountryCode=rstemp(21)
			pcv_shippingZip=rstemp(22)
			pcv_orderStatus=rstemp(23)
			pcv_shippingPhone=rstemp(24)
			pcv_idAffiliate=rstemp(25)
			pcv_processDate=rstemp(26)
			
			query="SELECT pcPackageInfo_ShippedDate FROM pcPackageInfo WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			pcv_shipDate=""
			if not rsStr.eof then
			do while not rsStr.eof
				tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
				if scDateFrmt="DD/MM/YY" then
					tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
				end if
				if  pcv_shipDate<>"" then
					pcv_shipDate=pcv_shipDate & vbcrlf
				end if
				pcv_shipDate=pcv_shipDate & tmp_processDate
				rsStr.MoveNext
			loop
			else
			query="SELECT shipDate FROM orders WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			if not rsStr.eof then
				pcv_shipDate=rsStr("shipDate")
				if scDateFrmt="DD/MM/YY" then
					pcv_shipDate=(day(pcv_shipDate)&"/"&month(pcv_shipDate)&"/"&year(pcv_shipDate))
				end if
			end if
			end if
			set rsStr=nothing
			
			pcv_shipVia=rstemp(28)
			pcv_trackingNum=rstemp(29)
			pcv_affiliatePay=rstemp(30)
			pcv_returnDate=rstemp(31)
			pcv_returnReason=rstemp(32)
			pcv_iRewardPoints=rstemp(33)
			if pcv_iRewardPoints<>"" then
			else
			pcv_iRewardPoints="0"
			end if
			pcv_ShippingFullName=rstemp(34)
			pcv_iRewardValue=rstemp(35)
			pcv_iRewardRefId=rstemp(36)
			pcv_iRewardPointsRef=rstemp(37)
			pcv_iRewardPointsCustAccrued=rstemp(38)
			pcv_IDRefer=rstemp(39)
			if pcv_IDRefer<>"" then
			else
			pcv_IDRefer="0"
			end if
			pcv_address2=rstemp(40)
			pcv_shippingCompany=rstemp(41)
			pcv_shippingAddress2=rstemp(42)
			pcv_taxDetails=rstemp(43)
			pcv_adminComments=rstemp(44)
			pcv_rmaCredit=rstemp(45)
			if pcv_rmaCredit<>"" then
			else
			pcv_rmaCredit="0"
			end if
			pcv_DPs=rstemp(46)
			pcv_gwAuthCode=rstemp(47)
			pcv_gwTransId=rstemp(48)
			pcv_paymentCode=rstemp(49)
			pcv_SRF=rstemp(50)
			pcv_ordShiptype=rstemp(51)
			if pcv_ordShiptype<>"" then
			else
			pcv_ordShiptype="0"
			end if
			pcv_ordPackageNum=rstemp(52)
			if pcv_ordPackageNum<>"" then
			else
			pcv_ordPackageNum="0"
			end if
			pcv_ord_DeliveryDate=rstemp(53)
			pcv_ord_OrderName=rstemp(54)
			pcv_ord_VAT=rstemp(55)
			if pcv_ord_VAT<>"" then
			else
			pcv_ord_VAT="0"
			end if

			pcv_pcOrd_CatDiscounts=rstemp(56)
			if pcv_pcOrd_CatDiscounts<>"" then
			else
			pcv_pcOrd_CatDiscounts="0"
			end if
			
			set StringBuilderObj = new StringBuilder
		 
			'***** Column 1 *****
			StringBuilderObj.append chr(34) & (int(pcv_idorder+scPre)) & chr(34) & ","
			
			'***** Column 2 *****
			StringBuilderObj.append chr(34) & replace(pcv_ord_OrderName,chr(34),chr(34) & chr(34)) & chr(34) & ","
			
			'***** Column 3 *****
			dtOrderDate=pcv_orderDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			StringBuilderObj.append chr(34) & dtOrderDate & chr(34) & ","
			
			'***** Column 4 *****
			if pcv_orderStatus="2" then
			strOStatus="Pending"
			end if
			if pcv_orderStatus="3" then
			strOStatus="Processed"
			end if
			if pcv_orderStatus="4" then
			strOStatus="Shipped"
			end if
			StringBuilderObj.append chr(34) & strOStatus & chr(34) & ","
			
			'***** Column 5 *****
			OrdDetails=replace(GetProductCFs(pcv_idOrder,pcv_details)," ||",""&scCurSign&"")
			StringBuilderObj.append chr(34) & replace(OrdDetails,chr(34),chr(34) & chr(34)) & chr(34) & ","
			
			'***** Column 6 *****
			dtOrderDate=pcv_processDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if

			StringBuilderObj.append chr(34) & dtOrderDate & chr(34) & ","
			
			'***** Column 7 *****
			query="select name,lastname,customerCompany,phone FROM customers where idcustomer=" & pcv_idCustomer
			set rs=connTemp.execute(query)
			
			pcv_CustName=rs("name") & " " & rs("lastname") & " (#" & pcv_idcustomer & ")"
			pcv_CustCompany=rs("customerCompany")
			pcv_CustPhone=rs("phone")
			tmpS1=""
			tmpS1=pcv_CustName & vbcrlf
			if trim(pcv_CustCompany)<>"" then
				tmpS1=tmpS1 & "Company: " & pcv_CustCompany & vbcrlf
			end if
			tmpS1=tmpS1 & pcv_address & vbcrlf
			if trim(pcv_address2)<>"" then
				tmpS1=tmpS1 & "Address 2: " & pcv_address2 & vbcrlf
			end if
			tmpS1=tmpS1 & pcv_city &", " & pcv_statecode & pcv_state & " " & pcv_zip & ", " & pcv_countryCode & vbcrlf
			if pcv_CustPhone<>"" then
				tmpS1=tmpS1 & "Phone: " & pcv_CustPhone
			end if
			StringBuilderObj.append chr(34) & replace(tmpS1,chr(34),chr(34) & chr(34)) & chr(34) & ","
			
			'***** Column 8 *****
			tmpS1=""
			if pcv_ShippingFullName<>"" then
				tmpS1=tmpS1 & "Shipping Name: " & pcv_ShippingFullName & vbcrlf
			end if
			if pcv_shippingCompany<>"" then
				tmpS1=tmpS1 & "Company: " & pcv_shippingCompany & vbcrlf
			end if
			tmpS1=tmpS1 & pcv_shippingAddress & vbcrlf
			if pcv_shippingAddress2<>"" then
				tmpS1=tmpS1 & "Address 2: " & pcv_shippingAddress2 & vbcrlf
			end if
			tmpS1=tmpS1 & pcv_shippingCity &", " & pcv_shippingStatecode & pcv_shippingState & " " & pcv_shippingZip & ", " & pcv_shippingCountryCode & vbcrlf
			if trim(pcv_shippingPhone)<>"" then
				tmpS1=tmpS1 & "Phone: " & pcv_shippingPhone
			end if
			
			StringBuilderObj.append chr(34) & replace(tmpS1,chr(34),chr(34) & chr(34)) & chr(34) & ","
			
			'***** Column 9 *****
			ShipDetails=""
			if pcv_SRF="1" then
			ShipDetails="Shipping charges to be determined."
			else
			if instr(pcv_shipmentDetails,",")=0 then
			ShipDetails=pcv_shipmentDetails
			end if			
			end if
			
			ShipFees=""
			HandlingFees=""
			if ShipDetails="" then
			pcv_Ship=split(pcv_shipmentDetails,",")
			ShipDetails=pcv_Ship(1)
			ShipFees=pcv_Ship(2)
			HandlingFees=pcv_Ship(3)
			end if
			
			StringBuilderObj.append chr(34) & "Shipping Method: " & ShipDetails & vbcrlf
			if pcv_ordShiptype="0" then
			StringBuilderObj.append "Shipping Type: Residential" & vbcrlf
			end if
			if pcv_ordShiptype="1" then
			StringBuilderObj.append "Shipping Type: Commercial" & vbcrlf
			end if
			if ShipFees>"0" then
			StringBuilderObj.append "Fees: " & scCurSign & money(ShipFees) & vbcrlf
			end if
			if HandlingFees>"0" then
			StringBuilderObj.append "Handling Fees: " & scCurSign & money(HandlingFees) & vbcrlf
			end if
			if pcv_ShipVia<>"" then
			StringBuilderObj.append "Shipped Via: " & pcv_ShipVia & vbcrlf
			end if
			if pcv_trackingNum<>"" then
			StringBuilderObj.append "Number of packages: " & pcv_trackingNum & vbcrlf
			end if
			if pcv_shipDate<>"" then
				dtOrderDate=pcv_shipDate
				StringBuilderObj.append "Date Shipped: " & dtOrderDate & vbcrlf
			end if
			if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			StringBuilderObj.append "Delivery Date: " & dtOrderDate
			end if
			
			StringBuilderObj.append chr(34) & ","
			
			'***** Column 10 *****
			StringBuilderObj.append chr(34)
			if instr(pcv_paymentDetails,"||")>0 then
				pcv_PayArray=split(pcv_paymentDetails,"||")
				StringBuilderObj.append "Payment Method: " & trim(pcv_PayArray(0)) & vbcrlf
				if trim(pcv_PayArray(1))<>"" then
					if IsNumeric(trim(pcv_PayArray(1))) then
						PayFees=cdbl(trim(pcv_PayArray(1)))
						if PayFees>0 then
							StringBuilderObj.append "Fees: " & scCurSign & money(PayFees) & vbcrlf
						end if
					end if
				end if
			else
				StringBuilderObj.append "Payment Details: " & pcv_paymentDetails & vbcrlf
			end if
			if pcv_paymentCode<>"" then
			StringBuilderObj.append "Payment Gateway: " & pcv_paymentCode & vbcrlf
			end if
			if pcv_gwTransId<>"" then
			StringBuilderObj.append "Transaction ID: " & pcv_gwTransId & vbcrlf
			end if
			if pcv_gwAuthCode<>"" then
			StringBuilderObj.append "Authorization Code: " & pcv_gwAuthCode
			end if
			StringBuilderObj.append chr(34) & ","
			
			'***** Column 11 *****
			StringBuilderObj.append chr(34)
			if instr(pcv_discountDetails,"- ||")>0 then
				tmpDArr=split(pcv_discountDetails,"- ||")
				For m=lbound(tmpDArr)+1 to ubound(tmpDArr)
					if instr(tmpDArr(m),",")>0 then
						tmpDArr(m)=replace(tmpDArr(m),",","mmmmm",1,1)
					end if
				Next
				pcv_discountDetails=Join(tmpDArr,"- ||")
				pcv_DisArray1=split(pcv_discountDetails,"mmmmm")
				For m=lbound(pcv_DisArray1) to ubound(pcv_DisArray1)
					if pcv_DisArray1(m)<>"" then
						pcv_DisArray=split(pcv_DisArray1(m),"- ||")
						DisAmount=cdbl(trim(pcv_DisArray(1)))
						StringBuilderObj.append "Discount Name: " & replace(trim(pcv_DisArray(0)),chr(34),chr(34) & chr(34)) & vbcrlf
						if DisAmount<>0 then
							StringBuilderObj.append "Amount: -" & scCurSign & money(DisAmount) & vbcrlf 
						end if
					end if
				Next

			end if
			
			if pcv_pcOrd_CatDiscounts<>"" then
			if cdbl(pcv_pcOrd_CatDiscounts)<>0 then
			StringBuilderObj.append "Discount by Categories: -" & scCurSign & money(pcv_pcOrd_CatDiscounts)
			end if
			end if
			StringBuilderObj.append chr(34) & ","
			
			'***** Column 12 *****
			StringBuilderObj.append chr(34) & "Tax Amount: " & scCurSign & money(pcv_taxAmount) & vbcrlf
			IF cdbl(pcv_taxAmount)>0 then
				if instr(pcv_taxDetails,",")>0 then
				TaxArray=split(pcv_taxDetails,",")
				For m=0 to (ubound(TaxArray)-1)
					tmpTax=split(TaxArray(m),"|")
					StringBuilderObj.append ucase(tmpTax(0)) & " - Amount: " & scCurSign & money(tmpTax(1)) & vbcrlf
				Next
				end if
			END IF
			StringBuilderObj.append chr(34) & ","
			
			'***** Column 13 *****
			StringBuilderObj.append chr(34) & scCurSign & money(pcv_ord_VAT) & chr(34) & ","
			
			'***** Column 14 *****
			StringBuilderObj.append chr(34)
			if pcv_rmaCredit>"0" then
			StringBuilderObj.append "-" & scCurSign & money(pcv_rmaCredit)
			end if
			StringBuilderObj.append chr(34) & ","
			
			'***** Column 15 *****
			StringBuilderObj.append chr(34)
			if pcv_iRewardPoints>"0" then
			StringBuilderObj.append scCurSign & money(pcv_iRewardPoints)
			end if
			StringBuilderObj.append chr(34) & ","

			
			'***** Column 16 *****
			StringBuilderObj.append chr(34)
			if pcv_idAffiliate>"1" then
			
			query="select affiliateName,commission FROM affiliates where idAffiliate=" & pcv_idAffiliate
			set rs=connTemp.execute(query)
			if not rs.eof then
				StringBuilderObj.append replace(rs("affiliateName"),chr(34),chr(34) & chr(34)) & " (#" & pcv_idAffiliate & ")" & vbcrlf
			
				if pcv_affiliatePay>"0" then
					StringBuilderObj.append "Amount: " & scCurSign & money(pcvTotal*(rs("commission")/100))
				end if
			end if
			
			end if
			StringBuilderObj.append chr(34) & ","
			
			'***** Column 17 *****
			StringBuilderObj.append chr(34)
			if pcv_idRefer>"0" then
			
			query="select Name FROM Referrer where IdRefer=" & pcv_idRefer
			set rs=connTemp.execute(query)
			
			StringBuilderObj.append replace(rs("Name"),chr(34),chr(34) & chr(34)) & " (#" & pcv_idRefer & ")"
			
			end if
			StringBuilderObj.append chr(34) & ","
			
			'***** Column 18 *****
			StringBuilderObj.append chr(34) & scCurSign & money(pcv_total) & chr(34) & ","
			
			'***** Column 19 *****
			StringBuilderObj.append chr(34) 
			if pcv_comments<>"" then
			StringBuilderObj.append replace(pcv_comments,chr(34),chr(34) & chr(34))
			end if
			StringBuilderObj.append chr(34) & ","

			
			'***** Column 20 *****
			StringBuilderObj.append chr(34)
			if pcv_adminComments<>"" then
			StringBuilderObj.append replace(pcv_adminComments,chr(34),chr(34) & chr(34))
			end if
			StringBuilderObj.append chr(34) & ","
			
			a.Write(StringBuilderObj.toString())
			a.Writeline()
			set StringBuilderObj = nothing
			rstemp.moveNext
		loop
		set rstemp=nothing
		response.redirect "getFile.asp?file="& strFile &"&Type=csv"
	End if
End Function


Function CreateXlsFile()

	Dim xlWorkSheet					' Excel Worksheet object
	Dim xlApplication 
				
	Set xlApplication=CreateObject("Excel.Application")
	xlApplication.Visible=False
	xlApplication.Workbooks.Add
	Set xlWorksheet=xlApplication.Worksheets(1)

	xlWorksheet.Cells(1,1).Value="Order #"
	xlWorksheet.Cells(1,1).Interior.ColorIndex=6
	xlWorksheet.Cells(1,2).Value="Order Name"
	xlWorksheet.Cells(1,2).Interior.ColorIndex=6
	xlWorksheet.Cells(1,3).Value="Order Date"
	xlWorksheet.Cells(1,3).Interior.ColorIndex=6
	xlWorksheet.Cells(1,4).Value="Order Status"
	xlWorksheet.Cells(1,4).Interior.ColorIndex=6
	xlWorksheet.Cells(1,5).Value="Order Details"
	xlWorksheet.Cells(1,5).Interior.ColorIndex=6
	xlWorksheet.Cells(1,6).Value="Processed Date"
	xlWorksheet.Cells(1,6).Interior.ColorIndex=6
	xlWorksheet.Cells(1,7).Value="Customer"
	xlWorksheet.Cells(1,7).Interior.ColorIndex=6
	xlWorksheet.Cells(1,8).Value="Shipment Address"
	xlWorksheet.Cells(1,8).Interior.ColorIndex=6
	xlWorksheet.Cells(1,9).Value="Shipment Details"
	xlWorksheet.Cells(1,9).Interior.ColorIndex=6
	xlWorksheet.Cells(1,10).Value="Payment Details"
	xlWorksheet.Cells(1,10).Interior.ColorIndex=6
	xlWorksheet.Cells(1,11).Value="Discount Details"
	xlWorksheet.Cells(1,11).Interior.ColorIndex=6
	xlWorksheet.Cells(1,12).Value="Tax Details"
	xlWorksheet.Cells(1,12).Interior.ColorIndex=6
	xlWorksheet.Cells(1,13).Value="VAT"
	xlWorksheet.Cells(1,13).Interior.ColorIndex=6
	xlWorksheet.Cells(1,14).Value="Credit"
	xlWorksheet.Cells(1,14).Interior.ColorIndex=6
	xlWorksheet.Cells(1,15).Value="Reward Points"
	xlWorksheet.Cells(1,15).Interior.ColorIndex=6
	xlWorksheet.Cells(1,16).Value="Affiliate"
	xlWorksheet.Cells(1,16).Interior.ColorIndex=6
	xlWorksheet.Cells(1,17).Value="Referrer"
	xlWorksheet.Cells(1,17).Interior.ColorIndex=6
	xlWorksheet.Cells(1,18).Value="Order Total"
	xlWorksheet.Cells(1,18).Interior.ColorIndex=6
	xlWorksheet.Cells(1,19).Value="Customer Comments"
	xlWorksheet.Cells(1,19).Interior.ColorIndex=6
	xlWorksheet.Cells(1,20).Value="Admin Comments"
	xlWorksheet.Cells(1,20).Interior.ColorIndex=6
			
	iRow=2
	If (DataEmpty=0) Then
		
	do until rstemp.eof
		 
pcv_idOrder=rstemp(0)
pcv_orderDate=rstemp(1)
pcv_idCustomer=rstemp(2)
pcv_details=rstemp(3)
pcv_total=rstemp(4)
if pcv_total<>"" then
else
pcv_total="0"
end if
pcv_address=rstemp(5)
pcv_zip=rstemp(6)
pcv_stateCode=rstemp(7)
pcv_state=rstemp(8)
pcv_city=rstemp(9)
pcv_countryCode=rstemp(10)
pcv_comments=rstemp(11)
pcv_taxAmount=rstemp(12)
if pcv_taxAmount<>"" then
else
pcv_taxAmount="0"
end if
pcv_shipmentDetails=rstemp(13)
pcv_paymentDetails=rstemp(14)
pcv_discountDetails=rstemp(15)
pcv_randomNumber=rstemp(16)
pcv_shippingAddress=rstemp(17)
pcv_shippingStateCode=rstemp(18)
pcv_shippingState=rstemp(19)
pcv_shippingCity=rstemp(20)
pcv_shippingCountryCode=rstemp(21)
pcv_shippingZip=rstemp(22)
pcv_orderStatus=rstemp(23)
pcv_shippingPhone=rstemp(24)
pcv_idAffiliate=rstemp(25)
pcv_processDate=rstemp(26)

			query="SELECT pcPackageInfo_ShippedDate FROM pcPackageInfo WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			pcv_shipDate=""
			if not rsStr.eof then
			do while not rsStr.eof
				tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
				if scDateFrmt="DD/MM/YY" then
					tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
				end if
				if  pcv_shipDate<>"" then
					pcv_shipDate=pcv_shipDate & vbcrlf
				end if
				pcv_shipDate=pcv_shipDate & tmp_processDate
				rsStr.MoveNext
			loop
			else
			query="SELECT shipDate FROM orders WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			if not rsStr.eof then
				pcv_shipDate=rsStr("shipDate")
				if scDateFrmt="DD/MM/YY" then
					pcv_shipDate=(day(pcv_shipDate)&"/"&month(pcv_shipDate)&"/"&year(pcv_shipDate))
				end if
			end if
			end if
			set rsStr=nothing

pcv_shipVia=rstemp(28)
pcv_trackingNum=rstemp(29)
pcv_affiliatePay=rstemp(30)
pcv_returnDate=rstemp(31)
pcv_returnReason=rstemp(32)
pcv_iRewardPoints=rstemp(33)
if pcv_iRewardPoints<>"" then
else
pcv_iRewardPoints="0"
end if
pcv_ShippingFullName=rstemp(34)
pcv_iRewardValue=rstemp(35)
pcv_iRewardRefId=rstemp(36)
pcv_iRewardPointsRef=rstemp(37)
pcv_iRewardPointsCustAccrued=rstemp(38)
pcv_IDRefer=rstemp(39)
if pcv_IDRefer<>"" then
else
pcv_IDRefer="0"
end if
pcv_address2=rstemp(40)
pcv_shippingCompany=rstemp(41)
pcv_shippingAddress2=rstemp(42)
pcv_taxDetails=rstemp(43)
pcv_adminComments=rstemp(44)
pcv_rmaCredit=rstemp(45)
if pcv_rmaCredit<>"" then
else
pcv_rmaCredit="0"
end if
pcv_DPs=rstemp(46)
pcv_gwAuthCode=rstemp(47)
pcv_gwTransId=rstemp(48)
pcv_paymentCode=rstemp(49)
pcv_SRF=rstemp(50)
pcv_ordShiptype=rstemp(51)
if pcv_ordShiptype<>"" then
else
pcv_ordShiptype="0"
end if
pcv_ordPackageNum=rstemp(52)
if pcv_ordPackageNum<>"" then
else
pcv_ordPackageNum="0"
end if
pcv_ord_DeliveryDate=rstemp(53)
pcv_ord_OrderName=rstemp(54)
pcv_ord_VAT=rstemp(55)
if pcv_ord_VAT<>"" then
else
pcv_ord_VAT="0"
end if

pcv_pcOrd_CatDiscounts=rstemp(56)
if pcv_pcOrd_CatDiscounts<>"" then
else
pcv_pcOrd_CatDiscounts="0"
end if

			'***** Column 1 *****
			xlWorksheet.Cells(iRow,1).Value=int(pcv_idorder+scPre)
			
			'***** Column 2 *****
			xlWorksheet.Cells(iRow,2).Value=pcv_ord_OrderName
			
			'***** Column 3 *****
			dtOrderDate=pcv_orderDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			xlWorksheet.Cells(iRow,3).Value=dtOrderDate
			
			'***** Column 4 *****
			if pcv_orderStatus="2" then
			strOStatus="Pending"
			end if
			if pcv_orderStatus="3" then
			strOStatus="Processed"
			end if
			if pcv_orderStatus="4" then
			strOStatus="Shipped"
			end if
			xlWorksheet.Cells(iRow,4).Value=strOStatus
			
			'***** Column 5 *****
			OrdDetails=replace(GetProductCFs(pcv_idOrder,pcv_details)," ||",""&scCurSign&"")
			xlWorksheet.Cells(iRow,5).Value=OrdDetails
			
			'***** Column 6 *****
			dtOrderDate=pcv_processDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			'Response.Write(dtOrderDate)
			xlWorksheet.Cells(iRow,6).Value=dtOrderDate
			
			'***** Column 7 *****
			query="select name,lastname,customerCompany,phone FROM customers where idcustomer=" & pcv_idCustomer
			set rs=connTemp.execute(query)
			
			pcv_CustName=rs("name") & " " & rs("lastname") & " (#" & pcv_idcustomer & ")"
			pcv_CustCompany=rs("customerCompany")
			pcv_CustPhone=rs("phone")
			
			set StringBuilderObj = new StringBuilder			
			StringBuilderObj.append pcv_CustName & vbcrlf
			if trim(pcv_CustCompany)<>"" then
				StringBuilderObj.append  "Company: " & pcv_CustCompany & vbcrlf
			end if
			StringBuilderObj.append  pcv_address & vbcrlf
			if trim(pcv_address2)<>"" then
				StringBuilderObj.append  "Address 2: " & pcv_address2 & vbcrlf
			end if
			StringBuilderObj.append  pcv_city &", " & pcv_statecode & pcv_state & " " & pcv_zip & ", " & pcv_countryCode & vbcrlf
			if pcv_CustPhone<>"" then
				StringBuilderObj.append  "Phone: " & pcv_CustPhone
			end if
			xlWorksheet.Cells(iRow,7).Value=StringBuilderObj.toString()
			set StringBuilderObj = nothing
			
			'***** Column 8 *****
			set StringBuilderObj = new StringBuilder
			if pcv_ShippingFullName<>"" then
				StringBuilderObj.append  "Shipping Name: " & pcv_ShippingFullName & vbcrlf
			end if
			if pcv_shippingCompany<>"" then
				StringBuilderObj.append  "Company: " & pcv_shippingCompany & vbcrlf
			end if
			StringBuilderObj.append  pcv_shippingAddress & vbcrlf
			if pcv_shippingAddress2<>"" then
				StringBuilderObj.append  "Address 2: " & pcv_shippingAddress2 & vbcrlf
			end if
			StringBuilderObj.append  pcv_shippingCity &", " & pcv_shippingStatecode & pcv_shippingState & " " & pcv_shippingZip & ", " & pcv_shippingCountryCode & vbcrlf
			if trim(pcv_shippingPhone)<>"" then
				StringBuilderObj.append  "Phone: " & pcv_shippingPhone
			end if			
			xlWorksheet.Cells(iRow,8).Value=StringBuilderObj.toString()			
			set StringBuilderObj = nothing
			
			'***** Column 9 *****
			ShipDetails=""
			if pcv_SRF="1" then
				ShipDetails="Shipping charges to be determined."
			else
				if instr(pcv_shipmentDetails,",")=0 then
					ShipDetails=pcv_shipmentDetails
				end if			
			end if
			
			ShipFees=""
			HandlingFees=""
			if ShipDetails="" then
				pcv_Ship=split(pcv_shipmentDetails,",")
				ShipDetails=pcv_Ship(1)
				ShipFees=pcv_Ship(2)
				HandlingFees=pcv_Ship(3)
			end if
			
			set StringBuilderObj = new StringBuilder			
			StringBuilderObj.append "Shipping Method: " & ShipDetails & vbcrlf
			if pcv_ordShiptype="0" then
				StringBuilderObj.append "Shipping Type: Residential" & vbcrlf
			end if
			if pcv_ordShiptype="1" then
				StringBuilderObj.append "Shipping Type: Commercial" & vbcrlf
			end if
			if ShipFees>"0" then
				StringBuilderObj.append "Fees: " & scCurSign & money(ShipFees) & vbcrlf
			end if
			if HandlingFees>"0" then
				StringBuilderObj.append "Handling Fees: " & scCurSign & money(HandlingFees) & vbcrlf
			end if
			if pcv_ShipVia<>"" then
				StringBuilderObj.append "Shipped Via: " & pcv_ShipVia & vbcrlf
			end if
			if pcv_trackingNum<>"" then
				StringBuilderObj.append "Number of packages: " & pcv_trackingNum & vbcrlf
			end if
			if pcv_shipDate<>"" then
				dtOrderDate=pcv_shipDate
				StringBuilderObj.append "Date Shipped: " & dtOrderDate & vbcrlf
			end if
			if pcv_ord_DeliveryDate<>"" then
				dtOrderDate=pcv_ord_DeliveryDate
				if scDateFrmt="DD/MM/YY" then
					dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
				end if
				StringBuilderObj.append "Delivery Date: " & dtOrderDate
			end if			
			xlWorksheet.Cells(iRow,9).Value=StringBuilderObj.toString()			
			set StringBuilderObj = nothing
			
			'***** Column 10 *****
			set StringBuilderObj = new StringBuilder			
			if instr(pcv_paymentDetails,"||")>0 then
				pcv_PayArray=split(pcv_paymentDetails,"||")
				StringBuilderObj.append "Payment Method: " & trim(pcv_PayArray(0)) & vbcrlf
				if trim(pcv_PayArray(1))<>"" then
					if IsNumeric(trim(pcv_PayArray(1))) then
						PayFees=cdbl(trim(pcv_PayArray(1)))
						if PayFees>0 then
							StringBuilderObj.append "Fees: " & scCurSign & money(PayFees) & vbcrlf
						end if
					end if
				end if
			else
				StringBuilderObj.append "Payment Details: " & pcv_paymentDetails & vbcrlf
			end if
			if pcv_paymentCode<>"" then
				StringBuilderObj.append "Payment Gateway: " & pcv_paymentCode & vbcrlf
			end if
			if pcv_gwTransId<>"" then
				StringBuilderObj.append "Transaction ID: " & pcv_gwTransId & vbcrlf
			end if
			if pcv_gwAuthCode<>"" then
				StringBuilderObj.append "Authorization Code: " & pcv_gwAuthCode
			end if
			xlWorksheet.Cells(iRow,10).Value=StringBuilderObj.toString()
			set StringBuilderObj = nothing
			
			'***** Column 11 *****
			set StringBuilderObj = new StringBuilder			
			if instr(pcv_discountDetails,"- ||")>0 then			
				tmpDArr=split(pcv_discountDetails,"- ||")
				For m=lbound(tmpDArr)+1 to ubound(tmpDArr)
					if instr(tmpDArr(m),",")>0 then
						tmpDArr(m)=replace(tmpDArr(m),",","mmmmm",1,1)
					end if
				Next
				pcv_discountDetails=Join(tmpDArr,"- ||")
				pcv_DisArray1=split(pcv_discountDetails,"mmmmm")
				DisAmount=cdbl(trim(pcv_DisArray(1)))
				StringBuilderObj.append "Discount Name: " & trim(pcv_DisArray(0)) & vbcrlf
				if DisAmount<>0 then
					StringBuilderObj.append "Amount: -" & scCurSign & money(DisAmount) & vbcrlf 
				end if			
			end if			
			if pcv_pcOrd_CatDiscounts<>"" then
				if cdbl(pcv_pcOrd_CatDiscounts)<>0 then
					StringBuilderObj.append "Discount by Categories: -" & scCurSign & money(pcv_pcOrd_CatDiscounts)
				end if
			end if
			xlWorksheet.Cells(iRow,11).Value=StringBuilderObj.toString
			set StringBuilderObj = nothing
			
			'***** Column 12 *****
			set StringBuilderObj = new StringBuilder
			StringBuilderObj.append "Tax Amount: " & scCurSign & money(pcv_taxAmount) & vbcrlf
			IF cdbl(pcv_taxAmount)>0 then
				if instr(pcv_taxDetails,0)>0 then
				TaxArray=split(pcv_taxDetails,",")
				For m=0 to (ubound(TaxArray)-1)
					tmpTax=split(TaxArray(m),"|")
					StringBuilderObj.append ucase(tmpTax(0)) & " - Amount: " & scCurSign & money(tmpTax(1)) & vbcrlf
				Next
				end if
			END IF
			xlWorksheet.Cells(iRow,12).Value=StringBuilderObj.toString
			set StringBuilderObj = nothing
			
			'***** Column 13 *****
			xlWorksheet.Cells(iRow,13).Value=scCurSign & money(pcv_ord_VAT)
			
			'***** Column 14 *****
			tmpS=""
			if pcv_rmaCredit>"0" then
				tmpS = tmpS & "-" & scCurSign & money(pcv_rmaCredit)
			end if
			xlWorksheet.Cells(iRow,14).Value=tmpS
			
			'***** Column 15 *****
			tmpS=""
			if pcv_iRewardPoints>"0" then
				tmpS = tmpS & scCurSign & money(pcv_iRewardPoints)
			end if
			xlWorksheet.Cells(iRow,15).Value=tmpS

			
			'***** Column 16 *****
			tmpS=""
			if pcv_idAffiliate>"1" then			
				query="select affiliateName,commission FROM affiliates where idAffiliate=" & pcv_idAffiliate
				set rs=connTemp.execute(query)
				if not rs.eof then
					tmpS = tmpS & rs("affiliateName") & " (#" & pcv_idAffiliate & ")" & vbcrlf				
					if pcv_affiliatePay>"0" then
						tmpS = tmpS & "Amount: " & scCurSign & money(pcvTotal*(rs("commission")/100))
					end if
				end if			
			end if
			xlWorksheet.Cells(iRow,16).Value=tmpS
			
			'***** Column 17 *****
			tmpS=""			
			if pcv_idRefer>"0" then			
				query="select Name FROM Referrer where IdRefer=" & pcv_idRefer
				set rs=connTemp.execute(query)
				tmpS = rs("Name") & " (#" & pcv_idRefer & ")"			
			end if
			xlWorksheet.Cells(iRow,17).Value=tmpS
			
			'***** Column 18 *****
			xlWorksheet.Cells(iRow,18).Value=scCurSign & money(pcv_total)
			
			'***** Column 19 *****
			tmpS=""
			if pcv_comments<>"" then
				tmpS = pcv_comments
			end if
			xlWorksheet.Cells(iRow,19).Value=tmpS
			
			'***** Column 20 *****
			tmpS=""
			if pcv_adminComments<>"" then
				tmpS = pcv_adminComments
			end if
			xlWorksheet.Cells(iRow,20).Value=tmpS

			
		iRow=iRow + 1
	rstemp.moveNext
	loop
End If

	strFile=GenFileName()
	xlWorksheet.SaveAs Server.MapPath(".") & "\" & strFile & ".xls"
	xlApplication.Quit												' Close the Workbook
	Set xlWorksheet=Nothing
	Set xlApplication=Nothing	
	response.redirect "getFile.asp?file="& strFile &"&Type=xls"
	
End Function
%>