<%@ LANGUAGE="VBSCRIPT" %>
<%
Server.ScriptTimeout = 5400
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Orders Report" %>
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
Response.Expires=0

dim query, conntemp, rstemp
call opendb()
' Choose the records to display
chk_idOrder= request.form("chk_idOrder")
chk_orderDate= request.form("chk_orderDate")
chk_idCustomer= request.form("chk_idCustomer")
chk_details= request.form("chk_details")
chk_total= request.form("chk_total")
chk_address= request.form("chk_address")
chk_zip= request.form("chk_zip")
chk_stateCode= request.form("chk_stateCode")
chk_state= request.form("chk_state")
chk_city= request.form("chk_city=")
chk_countryCode= request.form("chk_countryCode")
chk_comments= request.form("chk_comments")
chk_taxAmount= request.form("chk_taxAmount")
chk_shipmentDetails= request.form("chk_shipmentDetails")
chk_paymentDetails= request.form("chk_paymentDetails")
chk_discountDetails= request.form("chk_discountDetails")
chk_randomNumber= request.form("chk_randomNumber")
chk_shippingAddress= request.form("chk_shippingAddress")
chk_shippingStateCode= request.form("chk_shippingStateCode")
chk_shippingState= request.form("chk_shippingState")
chk_shippingCity= request.form("chk_shippingCity")
chk_shippingCountryCode= request.form("chk_shippingCountryCode")
chk_shippingZip= request.form("chk_shippingZip")
chk_orderStatus= request.form("chk_orderStatus")
chk_shippingPhone= request.form("chk_shippingPhone")
chk_idAffiliate= request.form("chk_idAffiliate")
chk_processDate= request.form("chk_processDate")
chk_shipDate= request.form("chk_shipDate")
chk_shipVia= request.form("chk_shipVia")
chk_trackingNum= request.form("chk_trackingNum")
chk_affiliatePay= request.form("chk_affiliatePay")
chk_returnDate= request.form("chk_returnDate")
chk_returnReason= request.form("chk_returnReason")
chk_iRewardPoints= request.form("chk_iRewardPoints")
chk_ShippingFullName= request.form("chk_ShippingFullName")
chk_iRewardValue= request.form("chk_iRewardValue")
chk_iRewardRefId= request.form("chk_iRewardRefId")
chk_iRewardPointsRef= request.form("chk_iRewardPointsRef")
chk_iRewardPointsCustAccrued= request.form("chk_iRewardPointsCustAccrued")
chk_IDRefer= request.form("chk_IDRefer")
chk_ReferName=request.form("chk_ReferName")
chk_address2= request.form("chk_address2")
chk_shippingCompany= request.form("chk_shippingCompany")
chk_shippingAddress2= request.form("chk_shippingAddress2")
chk_taxDetails= request.form("chk_taxDetails")
chk_adminComments= request.form("chk_adminComments")
chk_rmaCredit= request.form("chk_rmaCredit")
chk_DPs= request.form("chk_DPs")
chk_gwAuthCode= request.form("chk_gwAuthCode")
chk_gwTransId= request.form("chk_gwTransId")
chk_paymentCode= request.form("chk_paymentCode")
chk_SRF= request.form("chk_SRF")
chk_ordShiptype= request.form("chk_ordShiptype")
chk_ordPackageNum= request.form("chk_ordPackageNum")
chk_ord_DeliveryDate= request.form("chk_ord_DeliveryDate")
chk_ord_OrderName= request.form("chk_ord_OrderName")
chk_ord_VAT= request.form("chk_ord_VAT")
chk_pcOrd_CatDiscounts= request.form("chk_pcOrd_CatDiscounts")
chk_CustomerDetails= request.form("chk_CustomerDetails")
chk_taxDetails= request.form("chk_taxDetails")
chk_paymentDetails= request.form("chk_paymentDetails")
chk_pcOrd_DiscountDetails= request.form("chk_pcOrd_DiscountDetails")
chk_pcOrd_GiftCertificates= request.form("chk_pcOrd_GiftCertificates")
chk_shipmentDetails= request.form("chk_shipmentDetails")
chk_AffiliateName= request.form("chk_AffiliateName")
chk_DSNotify=request.form("chk_DSNotify")
err.clear

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

If DateVar<>"" then
	if SQL_Format="1" then
		DateVar=Day(DateVar)&"/"&Month(DateVar)&"/"&Year(DateVar)
	end if

	if scDB="Access" then
		query1=" AND orders.orderDate >=#" & DateVar & "# "
	else
		query1=" AND orders.orderDate >='" & DateVar & "' "
	end if
else
	query1=""		
End If

If DateVar2<>"" then
	if SQL_Format="1" then
		DateVar2=Day(DateVar2)&"/"&Month(DateVar2)&"/"&Year(DateVar2)
	end if
	if scDB="Access" then
		query2=" AND orders.orderDate <=#" & DateVar2 & "# "
	else
		query2=" AND orders.orderDate <='" & DateVar2 & "' "
	end if
else
	query2=""
End If

Dim intIncludeAll, TempSQLall
intIncludeAll = Request.Form("includeAll")
	if intIncludeAll = "1" then
		TempSQLall ="WHERE (orders.orderStatus>0 AND orders.orderStatus<7) "
		else
		TempSQLall ="WHERE ((orders.orderStatus>1 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) "
	end if

if (DateVar<>"") then
	'Fields=61 
	query="SELECT idOrder,orderDate,idCustomer,0,total,address,zip,stateCode,state,city,countryCode,0,taxAmount,shipmentDetails,paymentDetails,discountDetails,randomNumber,shippingAddress,shippingStateCode,shippingState,shippingCity,shippingCountryCode,shippingZip,orderStatus,pcOrd_shippingPhone,idAffiliate,processDate,shipDate,shipVia,trackingNum,affiliatePay,returnDate,returnReason,iRewardPoints,ShippingFullName,iRewardValue,iRewardRefId,iRewardPointsRef,iRewardPointsCustAccrued,IDRefer,address2,shippingCompany,shippingAddress2,0,0,rmaCredit,DPs,gwAuthCode,gwTransId,paymentCode,SRF,ordShiptype,ordPackageNum,ord_DeliveryDate,ord_OrderName,ord_VAT,pcOrd_CatDiscounts,pcOrd_GCDetails,details,comments,adminComments,taxDetails FROM orders " & TempSQLall & query1 & query2 & " ORDER BY orders.orderDate ASC;"
else
	query="SELECT idOrder,orderDate,idCustomer,0,total,address,zip,stateCode,state,city,countryCode,0,taxAmount,shipmentDetails,paymentDetails,discountDetails,randomNumber,shippingAddress,shippingStateCode,shippingState,shippingCity,shippingCountryCode,shippingZip,orderStatus,pcOrd_shippingPhone,idAffiliate,processDate,shipDate,shipVia,trackingNum,affiliatePay,returnDate,returnReason,iRewardPoints,ShippingFullName,iRewardValue,iRewardRefId,iRewardPointsRef,iRewardPointsCustAccrued,IDRefer,address2,shippingCompany,shippingAddress2,0,0,rmaCredit,DPs,gwAuthCode,gwTransId,paymentCode,SRF,ordShiptype,ordPackageNum,ord_DeliveryDate,ord_OrderName,ord_VAT,pcOrd_CatDiscounts,pcOrd_GCDetails,details,comments,adminComments,taxDetails FROM orders " & TempSQLall & " ORDER BY orders.orderDate ASC;"
end if

set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query) 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
		<title>Orders Report from <%=DateVar%> to <%=DateVar2%></title>
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

Function GetDSNotifyData(tmpIDOrder,rtype)
Dim rsQ,queryQ,intcountQ,i,pcArrQ
Dim tmpStr1,tmpStr2
Dim tmpDSID,tmpIsSup,tmpIDPrd
	tmpStr1="***"
	tmpStr2=""
	
	queryQ="SELECT pcDropShipper_ID,pcPrdOrd_SentNotice,idProduct,pcPrdOrd_SentNotice FROM ProductsOrdered WHERE idorder=" & tmpIDOrder & " AND pcDropShipper_ID>0;"
	set rsQ=connTemp.execute(queryQ)
	
	if not rsQ.eof then
		pcArrQ=rsQ.getRows()
		intCountQ=ubound(pcArrQ,2)
		set rsQ=nothing
		
		For i=0 to intCountQ
			tmpDSID=pcArrQ(0,i)
			tmpIDPrd=pcArrQ(2,i)
			tmpIsSup=0
			if Instr(tmpStr1,"***" & tmpDSID & "***")=0 then
				tmpStr1=tmpStr1 & tmpDSID & "***"
				queryQ="SELECT pcDS_IsDropShipper FROM pcDropShippersSuppliers WHERE idproduct=" & tmpIDPrd & ";"
				set rsQ=connTemp.execute(queryQ)
				if not rsQ.eof then
					tmpIsSup=rsQ("pcDS_IsDropShipper")
					if IsNull(tmpIsSup) or tmpIsSup="" then
						tmpIsSup=0
					end if
				end if
				set rsQ=nothing
				if tmpIsSup=0 then
					queryQ="SELECT pcDropShipper_Company As DSName FROM pcDropShippers WHERE pcDropShipper_ID=" & tmpDSID & ";"
				else
					queryQ="SELECT pcSupplier_Company As DSName FROM pcSuppliers WHERE pcSupplier_ID=" & tmpDSID & ";"
				end if
				set rsQ=connTemp.execute(queryQ)
				pcArrQ(3,i)="Unknown"
				if not rsQ.eof then
					pcArrQ(3,i)=rsQ("DSName")
				end if
				set rsQ=nothing
				if tmpStr2<>"" then
					if rtype=1 then
						tmpStr2=tmpStr2 & "<br>"
					else
						tmpStr2=tmpStr2 & " || "
					end if
				end if
				tmpStr2=tmpStr2 & pcArrQ(3,i) & ": "
				if pcArrQ(1,i)="1" then
					tmpStr2=tmpStr2 & "YES"
				else
					tmpStr2=tmpStr2 & "NO"
				end if
			end if
		Next
	end if
	set rsQ=nothing
	
	GetDSNotifyData=tmpStr2
End Function

Function GetProductCFs(idOrder,tmpOrdDetails)
Dim tmp1,queryQ,rsQ,tmpArr,intCount,i,j
	tmp1=split(tmpOrdDetails,vbcrlf)
	queryQ="SELECT Products.sku,Products.Description,ProductsOrdered.xfdetails FROM Products INNER JOIN ProductsOrdered ON Products.idProduct=ProductsOrdered.idProduct WHERE ProductsOrdered.idOrder=" & idOrder & " AND (ProductsOrdered.xfdetails IS NOT NULL);"
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

Function GenHTML() %>
	<h1>Orders Report from <%=DateVar%> to <%=DateVar2%></h1>
	<table class="salesExport">
	<tr>
<%if chk_idOrder="1" then %>
	<th>Order ID</th>
<% End If%>
<%if chk_ord_OrderName="1" then %>
	<th>Order Name</th>
<% End If%>
<%if chk_orderDate="1" then %>
	<th>Order Date</th>
<% End If%>
<%if chk_idCustomer="1" then %>
	<th>Customer ID</th>
<% End If%>
<%if chk_CustomerDetails="1" then %>
	<th>Customer Details</th>
<% End If%>
<%if chk_details="1" then %>
	<th>Order Details</th>
<% End If%>
<%if chk_total="1" then %>
	<th>Order Total</th>
<% End If%>
<%if chk_processDate="1" then %>
	<th>Processed Date</th>
<% End If%>
<%if chk_ShippingFullName="1" then %>
	<th>Shipping Name</th>
<% End If%>
<%if chk_shippingCompany="1" then %>
	<th>Shipping Company</th>
<% End If%>
<%if chk_shippingAddress="1" then %>
	<th>Shipping Address</th>
<% End If%>
<%if chk_shippingAddress2="1" then %>
	<th>Shipping Address 2</th>
<% End If%>
<%if chk_shippingCity="1" then %>
	<th>Shipping City</th>
<% End If%>
<%if chk_shippingStateCode="1" then %>
	<th>Shipping State</th>
<% End If%>
<%if chk_shippingState="1" then %>
	<th>Shipping Province</th>
<% End If%>
<%if chk_shippingCountryCode="1" then %>
	<th>Shipping Country</th>
<% End If%>
<%if chk_shippingZip="1" then %>
	<th>Shipping Zip</th>
<% End If%>
<%if chk_shippingPhone="1" then %>
	<th>Shipping Phone</th>
<% End If%>
<%if chk_ShipmentDetails="1" then %>
	<th>Shipment Details</th>
<% End If%>
<%if chk_ordShiptype="1" then %>
	<th>Shipping Type</th>
<% End If%>
<%if chk_ordPackageNum="1" then %>
	<th>Number of packages</th>
<% End If%>
<%if chk_shipDate="1" then %>
	<th>Shipping Date</th>
<% End If%>
<%if chk_shipVia="1" then %>
	<th>Shipped Via</th>
<% End If%>
<%if chk_trackingNum="1" then %>
	<th>Tracking Number</th>
<% End If%>
<%if chk_ord_DeliveryDate="1" then %>
	<th>Delivery Date</th>
<% End If%>
<%if chk_orderStatus="1" then %>
	<th>Order Status</th>
<% End If%>
<%if chk_PaymentDetails="1" then %>
	<th>Payment Details</th>
<% End If%>
<%if chk_idAffiliate="1" then %>
	<th>Affiliate ID</th>
<% End If%>
<%if chk_AffiliateName="1" then %>
	<th>Affiliate Name</th>
<% End If%>
<%if chk_affiliatePay="1" then %>
	<th>Affiliate Payment</th>
<% End If%>
<%if chk_iRewardPoints="1" then %>
	<th><%=RewardsLabel%></th>
<% End If%>
<%if chk_iRewardPointsCustAccrued="1" then %>
	<th>Accrued <%=RewardsLabel%></th>
<% End If%>
<% if chk_IDRefer="1" then %>
	<th>Referrer ID</th>
<% End If%>
<% if chk_ReferName="1" then %>
	<th>Referrer Name</th>
<% End If%>
<%if chk_rmaCredit="1" then %>
	<th>RMA Credit</th>
<% End If%>
<%if chk_gwAuthCode="1" then %>
	<th>Authorization Code</th>
<% End If%>
<%if chk_gwTransId="1" then %>
	<th>Transaction ID</th>
<% End If%>
<%if chk_paymentCode="1" then %>
	<th>Payment Gateway</th>
<% End If%>
<%if chk_taxAmount="1" then %>
	<th>Tax Amount</th>
<% End If%>
<%if chk_taxDetails="1" then %>
	<th>Tax Details</th>
<% End If%>
<%if chk_ord_VAT="1" then %>
	<th>VAT</th>
<% End If%>
<%if chk_pcOrd_DiscountDetails="1" then %>
	<th>Discount Details</th>
<% End If%>
<%if chk_pcOrd_CatDiscounts="1" then %>
	<th>Categories Discounts</th>
<% End If%>
<%if chk_pcOrd_GiftCertificates="1" then %>
	<th>Redeemed Gift Certificates</th>
<% End If%>
<%if chk_comments="1" then %>
	<th>Customer Comments</th>
<% End If%>
<%if chk_adminComments="1" then %>
	<th>Admin Comments</th>
<% End If%>
<%if chk_returnDate="1" then %>
	<th>Return Date</th>
<% End If%>
<%if chk_returnReason="1" then %>
	<th>Return Reason</th>
<% End If%>
<%if chk_DSNotify="1" then%>
	<th>Drop-shipper Notifications</th>
<% End If%>
	</tr>
<%	if(rstemp.BOF=True and rstemp.EOF=True) then%>
	<tr>
		<td colspan="50">No records found</td>
	</tr>
	<%
	set rstemp=nothing
else
	pcArr=rstemp.getRows()
	set rstemp=nothing
	intCount=ubound(pcArr,2)
	For nk=0 to intCount
	pcv_idOrder=pcArr(0,nk)
	pcv_ShowID=scpre+int(pcv_idOrder)
	pcv_orderDate=pcArr(1,nk)
	pcv_idCustomer=pcArr(2,nk)
	pcv_details=pcArr(58,nk)
	pcv_total=pcArr(4,nk)
	if pcv_total<>"" then
	else
		pcv_total="0"
	end if
	pcv_address=pcArr(5,nk)
	pcv_zip=pcArr(6,nk)
	pcv_stateCode=pcArr(7,nk)
	pcv_state=pcArr(8,nk)
	pcv_city=pcArr(9,nk)
	pcv_countryCode=pcArr(10,nk)
	pcv_comments=pcArr(59,nk)
	pcv_taxAmount=pcArr(12,nk)
	if pcv_taxAmount<>"" then
	else
		pcv_taxAmount="0"
	end if
	pcv_shipmentDetails=pcArr(13,nk)
	pcv_paymentDetails=pcArr(14,nk)
	pcv_discountDetails=pcArr(15,nk)
	pcv_randomNumber=pcArr(16,nk)
	pcv_shippingAddress=pcArr(17,nk)
	pcv_shippingStateCode=pcArr(18,nk)
	pcv_shippingState=pcArr(19,nk)
	pcv_shippingCity=pcArr(20,nk)
	pcv_shippingCountryCode=pcArr(21,nk)
	pcv_shippingZip=pcArr(22,nk)
	pcv_orderStatus=pcArr(23,nk)
	pcv_shippingPhone=pcArr(24,nk)
	pcv_idAffiliate=pcArr(25,nk)
	pcv_processDate=pcArr(26,nk)

	if (chk_shipDate="1") OR (chk_ShipmentDetails="1") OR (chk_shipVia="1") OR (chk_trackingNum="1") then

			query="SELECT pcPackageInfo_ID,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber FROM pcPackageInfo WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			pcv_shipDate=""
			pcv_shipVia=""
			pcv_trackingNum=""
			tmp_HavePacks=0
			if not rsStr.eof then
			do while not rsStr.eof
				tmp_HavePacks=1
				tmp_packID=rsStr("pcPackageInfo_ID")
				tmp_packMethod=rsStr("pcPackageInfo_ShipMethod")
				tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
				if scDateFrmt="DD/MM/YY" then
					tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
				end if
				tmp_TrackingNumber=rsStr("pcPackageInfo_TrackingNumber")
				
				if  pcv_shipDate<>"" then
					pcv_shipDate=pcv_shipDate & "<br>"
				end if
				pcv_shipDate=pcv_shipDate & tmp_processDate
				
				if  pcv_shipVia<>"" then
					pcv_shipVia=pcv_shipVia & "<br>"
				end if
				pcv_shipVia=pcv_shipVia & "Package ID# " & tmp_packID & " ** " & tmp_packMethod & " ** " & tmp_processDate & " ** " & tmp_TrackingNumber
				
				if  pcv_trackingNum<>"" then
					pcv_trackingNum=pcv_trackingNum & "<br>"
				end if
				pcv_trackingNum=pcv_trackingNum & tmp_TrackingNumber
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
				pcv_shipVia=pcArr(28,nk)
				pcv_trackingNum=pcArr(29,nk)
			end if
			set rsStr=nothing
			
	end if

	pcv_affiliatePay=pcArr(30,nk)
	pcv_returnDate=pcArr(31,nk)
	pcv_returnReason=pcArr(32,nk)
	pcv_iRewardPoints=pcArr(33,nk)
	if pcv_iRewardPoints<>"" then
	else
		pcv_iRewardPoints="0"
	end if
	pcv_ShippingFullName=pcArr(34,nk)
	pcv_iRewardValue=pcArr(35,nk)
	pcv_iRewardRefId=pcArr(36,nk)
	pcv_iRewardPointsRef=pcArr(37,nk)
	pcv_iRewardPointsCustAccrued=pcArr(38,nk)
	pcv_IDRefer=pcArr(39,nk)
	if isNull(pcv_IDRefer) or pcv_IDRefer="" then
		pcv_IDRefer="0"
	end if
	pcv_address2=pcArr(40,nk)
	pcv_shippingCompany=pcArr(41,nk)
	pcv_shippingAddress2=pcArr(42,nk)
	pcv_taxDetails=pcArr(61,nk)
	pcv_adminComments=pcArr(60,nk)
	pcv_rmaCredit=pcArr(45,nk)
	if pcv_rmaCredit<>"" then
	else
		pcv_rmaCredit="0"
	end if
	pcv_DPs=pcArr(46,nk)
	pcv_gwAuthCode=pcArr(47,nk)
	pcv_gwTransId=pcArr(48,nk)
	pcv_paymentCode=pcArr(49,nk)
	pcv_SRF=pcArr(50,nk)
	pcv_ordShiptype=pcArr(51,nk)
	if pcv_ordShiptype<>"" then
	else
		pcv_ordShiptype="0"
	end if
	pcv_ordPackageNum=pcArr(52,nk)
	if pcv_ordPackageNum<>"" then
	else
		pcv_ordPackageNum="0"
	end if
	pcv_ord_DeliveryDate=pcArr(53,nk)
	pcv_ord_OrderName=pcArr(54,nk)
	pcv_ord_VAT=pcArr(55,nk)
	if pcv_ord_VAT<>"" then
	else
		pcv_ord_VAT="0"
	end if

	pcv_pcOrd_CatDiscounts=pcArr(56,nk)
	if pcv_pcOrd_CatDiscounts<>"" then
	else
		pcv_pcOrd_CatDiscounts="0"
	end if
	
	pcv_pcOrd_RedeemedGC=pcArr(57,nk)
	if pcv_pcOrd_RedeemedGC<>"" then
	else
		pcv_pcOrd_RedeemedGC="0"
	end if
		
	%>

<tr valign="top">
<%if chk_idOrder="1" then %>
	<td><%=pcv_ShowID%></td>
<% End If%>
<%if chk_ord_OrderName="1" then %>
	<td><%=pcv_ord_OrderName%></td>
<% End If%>
<%if chk_orderDate="1" then
dtOrderDate=pcv_orderDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if %>
	<td><%=dtOrderDate%></td>
<% End If%>
<%if chk_idCustomer="1" then %>
	<td><%=pcv_IDCustomer%></td>
<% End If%>
<%if chk_CustomerDetails="1" then
	mySQL="select name,lastname,customerCompany,phone,email FROM customers where idcustomer=" & pcv_idCustomer
	set rs=connTemp.execute(mySQL)
			
	pcv_CustName=rs("name") & " " & rs("lastname") & " (#" & pcv_idcustomer & ")"
	pcv_CustCompany=rs("customerCompany")
	pcv_CustPhone=rs("phone")
	pcv_CustEmail=rs("email")
	
	set rs=nothing %>
	<td>
	<%
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
	if pcv_CustEmail<>"" then
		Response.Write("<br>Email: " & pcv_CustEmail)
	end if
	%></td>
<% End If%>
<%if chk_details="1" then %>
	<td><%Response.Write(replace(replace(GetProductCFs(pcv_idOrder,pcv_details)," ||",""&scCurSign&""),vbcrlf,"<br>"))%></td>
<% End If%>
<%if chk_total="1" then %>
	<td><%Response.Write(scCurSign & money(pcv_total))%></td>
<% End If%>
<%if chk_processDate="1" then
dtOrderDate=pcv_processDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if
%>
	<td><% Response.Write(dtOrderDate)%></td>
<% End If%>
<%if chk_ShippingFullName="1" then %>
	<td><%=pcv_ShippingFullName%></td>
<% End If%>
<%if chk_shippingCompany="1" then %>
	<td><%=pcv_shippingCompany%></td>
<% End If%>
<%if chk_shippingAddress="1" then %>
	<td><%=pcv_shippingAddress%></td>
<% End If%>
<%if chk_shippingAddress2="1" then %>
	<td><%=pcv_shippingAddress2%></td>
<% End If%>
<%if chk_shippingCity="1" then %>
	<td><%=pcv_shippingCity%></td>
<% End If%>
<%if chk_shippingStateCode="1" then %>
	<td><%=pcv_shippingStateCode%></td>
<% End If%>
<%if chk_shippingState="1" then %>
	<td><%=pcv_shippingState%></td>
<% End If%>
<%if chk_shippingCountryCode="1" then %>
	<td><%=pcv_shippingCountryCode%></td>
<% End If%>
<%if chk_shippingZip="1" then %>
	<td><%=pcv_shippingZip%></td>
<% End If%>
<%if chk_shippingPhone="1" then %>
	<td><%=pcv_shippingPhone%></td>
<% End If%>
<%if chk_ShipmentDetails="1" then %>
	<td>
	<%
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
				if tmp_HavePacks=0 then
					Response.Write("Shipped Via: " & pcv_ShipVia & "<br>")
				end if
			end if
			if pcv_ordPackageNum<>"" then
			Response.Write("Number of packages: " & pcv_ordPackageNum & "<br>")
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
	
	%></td>
<% End If%>
<%if chk_ordShiptype="1" then %>
	<td>
	<%
	if pcv_ordShiptype="0" then
	Response.Write("Shipping Type: Residential <br>")
	end if
	if pcv_ordShiptype="1" then
	Response.Write("Shipping Type: Commercial <br>")
	end if
	%></font>
	&nbsp;</td>
<% End If%>
<%if chk_ordPackageNum="1" then %>
	<td><%=pcv_ordPackageNum%></td>
<% End If%>
<%if chk_shipDate="1" then %>
	<td>
	<%
	if pcv_shipDate<>"" then
			dtOrderDate=pcv_shipDate
			Response.Write(dtOrderDate)
	end if
	%></td>
<% End If%>
<%if chk_shipVia="1" then %>
	<td><%=pcv_ShipVia%></td>
<% End If%>
<%if chk_trackingNum="1" then %>
	<td><%=pcv_trackingNum%></td>
<% End If%>
<%if chk_ord_DeliveryDate="1" then %>
	<td>
	<%
	if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			Response.Write(dtOrderDate)
	end if
	%></td>
<% End If%>
<%if chk_orderStatus="1" then
					OrderStatusStr="Incomplete"		
					Select case cint(pcv_orderStatus)
					Case 2: OrderStatusStr="Pending"
					Case 3: OrderStatusStr="Processed"
					Case 4: OrderStatusStr="Shipped"
					Case 5: OrderStatusStr="Canceled"
					Case 6: OrderStatusStr="Return"
					Case 7: OrderStatusStr="Partially Shipped"
					Case 8: OrderStatusStr="Shipping"
					Case 10: OrderStatusStr="Delivered"
					Case 12: OrderStatusStr="Archived"	
					End Select %>
	<td><%=OrderStatusStr%></td>
<% End If%>
<%if chk_PaymentDetails="1" then %>
	<td>
	<%
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
	%></td>
<% End If%>
<%if chk_idAffiliate="1" then %>
	<td><%=pcv_idAffiliate%></td>
<% End If%>
<%if chk_AffiliateName="1" then %>
	<td>
	<%
	if pcv_idAffiliate>"1" then
			
			mySQL="select affiliateName,commission FROM affiliates where idAffiliate=" & pcv_idAffiliate
			set rs=connTemp.execute(mySQL)
			
			Response.Write(rs("affiliateName") & " (#" & pcv_idAffiliate & ")" & "<br>")
			
			set rs=nothing
			
	end if
	%></td>
<% End If%>
<%if chk_affiliatePay="1" then %>
	<td><%=scCurSign & money(pcv_affiliatePay)%></td>
<% End If%>
<%if chk_iRewardPoints="1" then %>
	<td><%=pcv_iRewardPoints%></td>
<% End If%>
<%if chk_iRewardPointsCustAccrued="1" then %>
	<td><%=pcv_iRewardPointsCustAccrued%></td>
<% End If%>
<%if chk_IDRefer="1" then %>
	<td><%=pcv_IDRefer%></td>
<% End If%>
<%if chk_ReferName="1" then
	if pcv_IDRefer<>"0" then
		query="SELECT name FROM REFERRER where IdRefer="&pcv_IDRefer
		set rsReferObj=Server.CreateObject("ADODB.RecordSet")
		set rsReferObj=conntemp.execute(query)
		if NOT rsReferObj.eof then
			pcv_ReferName=rsReferObj("name")
		else
			pcv_ReferName=""
		end if
		set rsReferObj=nothing
	else
		pcv_ReferName=""
	end if %>
	<td><%=pcv_ReferName%></td>
<% End If%>
<%if chk_rmaCredit="1" then %>
	<td><%=pcv_rmaCredit%></td>
<% End If%>
<%if chk_gwAuthCode="1" then %>
	<td><%=pcv_gwAuthCode%></td>
<% End If%>
<%if chk_gwTransId="1" then %>
	<td><%=pcv_gwTransId%></td>
<% End If%>
<%if chk_paymentCode="1" then %>
	<td><%=pcv_paymentCode%></td>
<% End If%>
<%if chk_taxAmount="1" then %>
	<td><%=scCurSign & money(pcv_taxAmount)%></td>
<% End If%>
<%if chk_taxDetails="1" then %>
	<td>
	<%
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
	%></td>
<% End If%>
<%if chk_ord_VAT="1" then %>
	<td><%=scCurSign & money(pcv_ord_VAT)%></td>
<% End If%>
<%if chk_pcOrd_DiscountDetails="1" then %>
	<td>
	<%
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
				Response.Write("Discount Name: " & replace(trim(pcv_DisArray(0)),chr(34),chr(34) & chr(34)) & "<br>")
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
	%></td>
<% End If%>
<%if chk_pcOrd_CatDiscounts="1" then %>
	<td><%="-" & scCurSign & money(pcv_pcOrd_CatDiscounts)%></td>
<% End If%>
<%if chk_pcOrd_GiftCertificates="1" then %>
    <td>
    <%
	if pcv_pcOrd_RedeemedGC<>"0" then
		GCArry=split(pcv_pcOrd_RedeemedGC,"|g|")
		intArryCnt=ubound(GCArry)
	
		for k=0 to intArryCnt
	
		if GCArry(k)<>"" then
			GCInfo = split(GCArry(k),"|s|")
			if GCInfo(2)="" OR IsNull(GCInfo(2)) then
				GCInfo(2)=0
			end if
			pGiftCode=GCInfo(0)
			pGiftUsed=GCInfo(2)
		query="SELECT products.IDProduct,products.Description FROM pcGCOrdered,Products WHERE products.idproduct=pcGCOrdered.pcGO_idproduct AND pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
		SET rsGC=server.CreateObject("ADODB.RecordSet")
		SET rsGC=connTemp.execute(query)
	
		if NOT rsGC.eof then
			pIdproduct=rsGC("idproduct")
			pName=rsGC("Description")
			pCode=pGiftCode
			if k>0 then response.write "<br>"
			%>
			GC Name: <%=pName%><br>
				<% query="SELECT pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status FROM pcGCOrdered WHERE pcGO_GcCode='" & pGiftCode & "'"
				SET rsGCObj=server.CreateObject("ADODB.RecordSet")
				SET rsGCObj=connTemp.execute(query)
				
				if NOT rsGCObj.eof then
					pcGO_GcCode=rsGCObj("pcGO_GcCode")
					pExpDate=rsGCObj("pcGO_ExpDate")
					pGCAmount=rsGCObj("pcGO_Amount")
					pGCStatus=rsGCObj("pcGO_Status")
					%>
					GC Code: <b><%=pcGO_GcCode%></b><br>
					Used for this order:&nbsp;<%=scCurSign & money(pGiftUsed)%><br>
					<% if cdbl(pGCAmount)<=0 then%>
						Completely redeemed.
					<% else %>
						Available Amount: <%=scCurSign & money(pGCAmount)%>
						<br>
						<% if year(pExpDate)="1900" then%>
							GC does not expire
						<%else
							if scDateFrmt="DD/MM/YY" then
								pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
							else
								pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
							end if %>
							Expiring: <%=pExpDate%>
						<%end if%>
						<br>
						<% if pGCStatus="1" then%>
							Status: Active
						<%else%>
							Status: Inactive
						<%end if%>
					<%end if%>
					<br>
				<%end if
				set rsGCObj=nothing
				%>
		<%
			end if
			set rsGC=nothing
		end if
		Next
	end if
	%>
	</td>
<%end if%>
<%if chk_comments="1" then %>
	<td><%=pcv_comments%></td>
<% End If%>
<%if chk_adminComments="1" then %>
	<td><%=pcv_admincomments%></td>
<% End If%>
<%if chk_returnDate="1" then %>
	<td><%=pcv_returnDate%></td>
<% End If%>
<%if chk_returnReason="1" then %>
	<td><%=pcv_returnReason%></td>
<% End If%>
<%if chk_DSNotify="1" then%>
	<td><%=GetDSNotifyData(pcv_idOrder,1)%></td>
<% End if%>

</tr>

	<%Next
End if
set rstemp=nothing%>
    <tr>
    	<td colspan="50">Report Created On: <%=now()%></td>
    </tr>
</table>
<%
closedb()
End Function

Function CreateCSVFile()
	strFile=GenFileName()   
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
	If Not rstemp.EOF Then
		set StringBuilderObj = new StringBuilder
if chk_idOrder="1" then 
StringBuilderObj.append chr(34) & "Order ID" & chr(34) & ","
End If
if chk_ord_OrderName="1" then 
StringBuilderObj.append chr(34) & "Order Name" & chr(34) & ","
End If
if chk_orderDate="1" then 
StringBuilderObj.append chr(34) & "Order Date" & chr(34) & ","
End If
if chk_idCustomer="1" then 
StringBuilderObj.append chr(34) & "Customer ID" & chr(34) & ","
End If
if chk_CustomerDetails="1" then 
StringBuilderObj.append chr(34) & "Customer Details" & chr(34) & ","
End If
if chk_details="1" then 
StringBuilderObj.append chr(34) & "Order Details" & chr(34) & ","
End If
if chk_total="1" then 
StringBuilderObj.append chr(34) & "Order Total" & chr(34) & ","
End If
if chk_processDate="1" then 
StringBuilderObj.append chr(34) & "Processed Date" & chr(34) & ","
End If
if chk_ShippingFullName="1" then 
StringBuilderObj.append chr(34) & "Shipping Name" & chr(34) & ","
End If
if chk_shippingCompany="1" then 
StringBuilderObj.append chr(34) & "Shipping Company" & chr(34) & ","
End If
if chk_shippingAddress="1" then 
StringBuilderObj.append chr(34) & "Shipping Address" & chr(34) & ","
End If
if chk_shippingAddress2="1" then 
StringBuilderObj.append chr(34) & "Shipping Address 2" & chr(34) & ","
End If
if chk_shippingCity="1" then 
StringBuilderObj.append chr(34) & "Shipping City" & chr(34) & ","
End If
if chk_shippingStateCode="1" then 
StringBuilderObj.append chr(34) & "Shipping State" & chr(34) & ","
End If
if chk_shippingState="1" then 
StringBuilderObj.append chr(34) & "Shipping Province" & chr(34) & ","
End If
if chk_shippingCountryCode="1" then 
StringBuilderObj.append chr(34) & "Shipping Country" & chr(34) & ","
End If
if chk_shippingZip="1" then 
StringBuilderObj.append chr(34) & "Shipping Zip" & chr(34) & ","
End If
if chk_shippingPhone="1" then 
StringBuilderObj.append chr(34) & "Shipping Phone" & chr(34) & ","
End If
if chk_ShipmentDetails="1" then 
StringBuilderObj.append chr(34) & "Shipment Details" & chr(34) & ","
End If
if chk_ordShiptype="1" then 
StringBuilderObj.append chr(34) & "Shipping Type" & chr(34) & ","
End If
if chk_ordPackageNum="1" then 
StringBuilderObj.append chr(34) & "Number of packages" & chr(34) & ","
End If
if chk_shipDate="1" then 
StringBuilderObj.append chr(34) & "Shipping Date" & chr(34) & ","
End If
if chk_shipVia="1" then 
StringBuilderObj.append chr(34) & "Shipped Via" & chr(34) & ","
End If
if chk_trackingNum="1" then 
StringBuilderObj.append chr(34) & "Tracking Number" & chr(34) & ","
End If
if chk_ord_DeliveryDate="1" then 
StringBuilderObj.append chr(34) & "Delivery Date" & chr(34) & ","
End If
if chk_orderStatus="1" then 
StringBuilderObj.append chr(34) & "Order Status" & chr(34) & ","
End If
if chk_PaymentDetails="1" then 
StringBuilderObj.append chr(34) & "Payment Details" & chr(34) & ","
End If
if chk_idAffiliate="1" then 
StringBuilderObj.append chr(34) & "Affiliate ID" & chr(34) & ","
End If
if chk_AffiliateName="1" then 
StringBuilderObj.append chr(34) & "Affiliate Name" & chr(34) & ","
End If
if chk_affiliatePay="1" then 
StringBuilderObj.append chr(34) & "Affiliate Payment" & chr(34) & ","
End If
if chk_iRewardPoints="1" then 
StringBuilderObj.append chr(34) & RewardsLabel & chr(34) & ","
End If
if chk_iRewardPointsCustAccrued="1" then 
StringBuilderObj.append chr(34) & "Accrued " & RewardsLabel & chr(34) & ","
End If
if chk_IDRefer="1" then 
StringBuilderObj.append chr(34) & "Referrer ID" & chr(34) & ","
End If
if chk_ReferName="1" then 
StringBuilderObj.append chr(34) & "Referrer Name" & chr(34) & ","
End If
if chk_rmaCredit="1" then 
StringBuilderObj.append chr(34) & "RMA Credit" & chr(34) & ","
End If
if chk_gwAuthCode="1" then 
StringBuilderObj.append chr(34) & "Authorization Code" & chr(34) & ","
End If
if chk_gwTransId="1" then 
StringBuilderObj.append chr(34) & "Transaction ID" & chr(34) & ","
End If
if chk_paymentCode="1" then 
StringBuilderObj.append chr(34) & "Payment Gateway" & chr(34) & ","
End If
if chk_taxAmount="1" then 
StringBuilderObj.append chr(34) & "Tax Amount" & chr(34) & ","
End If
if chk_taxDetails="1" then 
StringBuilderObj.append chr(34) & "Tax Details" & chr(34) & ","
End If
if chk_ord_VAT="1" then 
StringBuilderObj.append chr(34) & "VAT" & chr(34) & ","
End If
if chk_pcOrd_DiscountDetails="1" then 
StringBuilderObj.append chr(34) & "Discount Details" & chr(34) & ","
End If
if chk_pcOrd_CatDiscounts="1" then 
StringBuilderObj.append chr(34) & "Categories Discounts" & chr(34) & ","
End If
if chk_pcOrd_GiftCertificates="1" then 
StringBuilderObj.append chr(34) & "Redeemed Gift Certificates" & chr(34) & ","
End If
if chk_comments="1" then 
StringBuilderObj.append chr(34) & "Customer Comments" & chr(34) & ","
End If
if chk_adminComments="1" then 
StringBuilderObj.append chr(34) & "Admin Comments" & chr(34) & ","
End If
if chk_returnDate="1" then 
StringBuilderObj.append chr(34) & "Return Date" & chr(34) & ","
End If
if chk_returnReason="1" then 
StringBuilderObj.append chr(34) & "Return Reason" & chr(34) & ","
End If
if chk_DSNotify="1" then
StringBuilderObj.append chr(34) & "Drop-shipper Notifications" & chr(34) & ","
End if
a.WriteLine(StringBuilderObj.toString())
set StringBuilderObj = nothing

	pcArr=rstemp.getRows()
	set rstemp=nothing
	intCount=ubound(pcArr,2)
	For nk=0 to intCount
	pcv_idOrder=pcArr(0,nk)
	pcv_ShowID=scpre+int(pcv_idOrder)
	pcv_orderDate=pcArr(1,nk)
	pcv_idCustomer=pcArr(2,nk)
	pcv_details=pcArr(58,nk)
	pcv_total=pcArr(4,nk)
	if pcv_total<>"" then
	else
		pcv_total="0"
	end if
	pcv_address=pcArr(5,nk)
	pcv_zip=pcArr(6,nk)
	pcv_stateCode=pcArr(7,nk)
	pcv_state=pcArr(8,nk)
	pcv_city=pcArr(9,nk)
	pcv_countryCode=pcArr(10,nk)
	pcv_comments=pcArr(59,nk)
	pcv_taxAmount=pcArr(12,nk)
	if pcv_taxAmount<>"" then
	else
		pcv_taxAmount="0"
	end if
	pcv_shipmentDetails=pcArr(13,nk)
	pcv_paymentDetails=pcArr(14,nk)
	pcv_discountDetails=pcArr(15,nk)
	pcv_randomNumber=pcArr(16,nk)
	pcv_shippingAddress=pcArr(17,nk)
	pcv_shippingStateCode=pcArr(18,nk)
	pcv_shippingState=pcArr(19,nk)
	pcv_shippingCity=pcArr(20,nk)
	pcv_shippingCountryCode=pcArr(21,nk)
	pcv_shippingZip=pcArr(22,nk)
	pcv_orderStatus=pcArr(23,nk)
	pcv_shippingPhone=pcArr(24,nk)
	pcv_idAffiliate=pcArr(25,nk)
	pcv_processDate=pcArr(26,nk)

	if (chk_shipDate="1") OR (chk_ShipmentDetails="1") OR (chk_shipVia="1") OR (chk_trackingNum="1") then

			query="SELECT pcPackageInfo_ID,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber FROM pcPackageInfo WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			pcv_shipDate=""
			pcv_shipVia=""
			pcv_trackingNum=""
			tmp_HavePacks=0
			if not rsStr.eof then
			do while not rsStr.eof
				tmp_HavePacks=1
				tmp_packID=rsStr("pcPackageInfo_ID")
				tmp_packMethod=rsStr("pcPackageInfo_ShipMethod")
				tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
				if scDateFrmt="DD/MM/YY" then
					tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
				end if
				tmp_TrackingNumber=rsStr("pcPackageInfo_TrackingNumber")
				
				if  pcv_shipDate<>"" then
					pcv_shipDate=pcv_shipDate & vbcrlf
				end if
				pcv_shipDate=pcv_shipDate & tmp_processDate
				
				if  pcv_shipVia<>"" then
					pcv_shipVia=pcv_shipVia & vbcrlf
				end if
				pcv_shipVia=pcv_shipVia & "Package ID# " & tmp_packID & " ** " & tmp_packMethod & " ** " & tmp_processDate & " ** " & tmp_TrackingNumber
				
				if  pcv_trackingNum<>"" then
					pcv_trackingNum=pcv_trackingNum & vbcrlf
				end if
				pcv_trackingNum=pcv_trackingNum & tmp_TrackingNumber
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
				pcv_shipVia=pcArr(28,nk)
				pcv_trackingNum=pcArr(29,nk)
			end if
			set rsStr=nothing
			
	end if
			
	pcv_affiliatePay=pcArr(30,nk)
	pcv_returnDate=pcArr(31,nk)
	pcv_returnReason=pcArr(32,nk)
	pcv_iRewardPoints=pcArr(33,nk)
	if pcv_iRewardPoints<>"" then
	else
		pcv_iRewardPoints="0"
	end if
	pcv_ShippingFullName=pcArr(34,nk)
	pcv_iRewardValue=pcArr(35,nk)
	pcv_iRewardRefId=pcArr(36,nk)
	pcv_iRewardPointsRef=pcArr(37,nk)
	pcv_iRewardPointsCustAccrued=pcArr(38,nk)
	pcv_IDRefer=pcArr(39,nk)
	if isNull(pcv_IDRefer) or pcv_IDRefer="" then
		pcv_IDRefer="0"
	end if
	pcv_address2=pcArr(40,nk)
	pcv_shippingCompany=pcArr(41,nk)
	pcv_shippingAddress2=pcArr(42,nk)
	pcv_taxDetails=pcArr(61,nk)
	pcv_adminComments=pcArr(60,nk)
	pcv_rmaCredit=pcArr(45,nk)
	if pcv_rmaCredit<>"" then
	else
		pcv_rmaCredit="0"
	end if
	pcv_DPs=pcArr(46,nk)
	pcv_gwAuthCode=pcArr(47,nk)
	pcv_gwTransId=pcArr(48,nk)
	pcv_paymentCode=pcArr(49,nk)
	pcv_SRF=pcArr(50,nk)
	pcv_ordShiptype=pcArr(51,nk)
	if pcv_ordShiptype<>"" then
	else
		pcv_ordShiptype="0"
	end if
	pcv_ordPackageNum=pcArr(52,nk)
	if pcv_ordPackageNum<>"" then
	else
		pcv_ordPackageNum="0"
	end if
	pcv_ord_DeliveryDate=pcArr(53,nk)
	pcv_ord_OrderName=pcArr(54,nk)
	pcv_ord_VAT=pcArr(55,nk)
	if pcv_ord_VAT<>"" then
	else
		pcv_ord_VAT="0"
	end if

	pcv_pcOrd_CatDiscounts=pcArr(56,nk)
	if pcv_pcOrd_CatDiscounts<>"" then
	else
		pcv_pcOrd_CatDiscounts="0"
	end if
	
	pcv_pcOrd_RedeemedGC=pcArr(57,nk)
	if pcv_pcOrd_RedeemedGC<>"" then
	else
		pcv_pcOrd_RedeemedGC="0"
	end if

set StringBuilderObj = new StringBuilder

if chk_idOrder="1" then 
StringBuilderObj.append chr(34)  &pcv_ShowID & chr(34) & ","
End If
if chk_ord_OrderName="1" then  
StringBuilderObj.append chr(34)  &pcv_ord_OrderName & chr(34) & ","
End If
if chk_orderDate="1" then
dtOrderDate=pcv_orderDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if
	StringBuilderObj.append chr(34)  & dtOrderDate & chr(34) & ","
End If
if chk_idCustomer="1" then 
StringBuilderObj.append chr(34)  &pcv_IDCustomer & chr(34) & ","
End If
if chk_CustomerDetails="1" then
	mySQL="select name,lastname,customerCompany,phone,email FROM customers where idcustomer=" & pcv_idCustomer
	set rs=connTemp.execute(mySQL)
			
	pcv_CustName=rs("name") & " " & rs("lastname") & " (#" & pcv_idcustomer & ")"
	pcv_CustCompany=rs("customerCompany")
	pcv_CustPhone=rs("phone")
	pcv_CustEmail=rs("email")
	
	set rs=nothing 
	
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
	if pcv_CustEmail<>"" then
		tmpS1=tmpS1 & vbcrlf & "Email: " & pcv_CustEmail
	end if
	StringBuilderObj.append chr(34) & replace(tmpS1,chr(34),chr(34) & chr(34)) & chr(34) & ","
End If
if chk_details="1" then 
	OrdDetails=replace(GetProductCFs(pcv_idOrder,pcv_details)," ||",""&scCurSign&"")
	StringBuilderObj.append chr(34) & replace(OrdDetails,chr(34),chr(34) & chr(34)) & chr(34) & ","
End If
if chk_total="1" then
	StringBuilderObj.append chr(34) & scCurSign & money(pcv_total) & chr(34) & ","
End If
if chk_processDate="1" then
dtOrderDate=pcv_processDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if
StringBuilderObj.append chr(34) & dtOrderDate & chr(34) & ","
End If
if chk_ShippingFullName="1" then  
StringBuilderObj.append chr(34)  &pcv_ShippingFullName & chr(34) & ","
End If
if chk_shippingCompany="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingCompany & chr(34) & ","
End If
if chk_shippingAddress="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingAddress & chr(34) & ","
End If
if chk_shippingAddress2="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingAddress2 & chr(34) & ","
End If
if chk_shippingCity="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingCity & chr(34) & ","
End If
if chk_shippingStateCode="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingStateCode & chr(34) & ","
End If
if chk_shippingState="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingState & chr(34) & ","
End If
if chk_shippingCountryCode="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingCountryCode & chr(34) & ","
End If
if chk_shippingZip="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingZip & chr(34) & ","
End If
if chk_shippingPhone="1" then  
StringBuilderObj.append chr(34)  &pcv_shippingPhone & chr(34) & ","
End If
if chk_ShipmentDetails="1" then 
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
			
			tmpS=""
			
			tmpS=tmpS & chr(34) & "Shipping Method: " & ShipDetails & vbcrlf
			if pcv_ordShiptype="0" then
			tmpS=tmpS & "Shipping Type: Residential" & vbcrlf
			end if
			if pcv_ordShiptype="1" then
			tmpS=tmpS & "Shipping Type: Commercial" & vbcrlf
			end if
			if ShipFees>"0" then
			tmpS=tmpS & "Fees: " & scCurSign & money(ShipFees) & vbcrlf
			end if
			if HandlingFees>"0" then
			tmpS=tmpS & "Handling Fees: " & scCurSign & money(HandlingFees) & vbcrlf
			end if
			if pcv_ShipVia<>"" then
				if tmp_HavePacks=0 then
					tmpS=tmpS & "Shipped Via: " & pcv_ShipVia & vbcrlf
				end if
			end if
			if pcv_ordPackageNum<>"" then
			tmpS=tmpS & "Number of packages: " & pcv_ordPackageNum & vbcrlf
			end if
			if pcv_shipDate<>"" then
				dtOrderDate=pcv_shipDate
				tmpS=tmpS & "Date Shipped: " & dtOrderDate & vbcrlf
			end if
			if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			tmpS=tmpS & "Delivery Date: " & dtOrderDate
			end if
			
			tmpS=tmpS & chr(34) & ","
			StringBuilderObj.append tmpS
End If
if chk_ordShiptype="1" then
			tmpS=""
			if pcv_ordShiptype="0" then
			tmpS=tmpS & "Shipping Type: Residential"
			end if
			if pcv_ordShiptype="1" then
			tmpS=tmpS & "Shipping Type: Commercial"
			end if
			StringBuilderObj.append chr(34) & tmpS & chr(34) & ","
End If
if chk_ordPackageNum="1" then 
StringBuilderObj.append chr(34)  &pcv_ordPackageNum & chr(34) & ","
End If
if chk_shipDate="1" then
	StringBuilderObj.append chr(34)
	if pcv_shipDate<>"" then
			dtOrderDate=pcv_shipDate
			StringBuilderObj.append dtOrderDate
	end if
	StringBuilderObj.append chr(34) & ","
End If
if chk_shipVia="1" then  
StringBuilderObj.append chr(34)  &pcv_ShipVia & chr(34) & ","
End If
if chk_trackingNum="1" then 
StringBuilderObj.append chr(34)  &pcv_trackingNum & chr(34) & ","
End If
if chk_ord_DeliveryDate="1" then
	StringBuilderObj.append chr(34)
	if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			StringBuilderObj.append dtOrderDate
	end if
	StringBuilderObj.append chr(34) & ","
End If
if chk_orderStatus="1" then
					OrderStatusStr="Incomplete"		
					Select case cint(pcv_orderStatus)
					Case 2: OrderStatusStr="Pending"
					Case 3: OrderStatusStr="Processed"
					Case 4: OrderStatusStr="Shipped"
					Case 5: OrderStatusStr="Canceled"
					Case 6: OrderStatusStr="Return"
					Case 7: OrderStatusStr="Partially Shipped"
					Case 8: OrderStatusStr="Shipping"
					Case 10: OrderStatusStr="Delivered"
					Case 12: OrderStatusStr="Archived"
					End Select
					StringBuilderObj.append chr(34) & OrderStatusStr & chr(34) & ","
End If
if chk_PaymentDetails="1" then
			tmpS=""
			tmpS=tmpS & chr(34)
			if instr(pcv_paymentDetails,"||")>0 then
				pcv_PayArray=split(pcv_paymentDetails,"||")
				tmpS=tmpS & "Payment Method: " & trim(pcv_PayArray(0)) & vbcrlf
				if trim(pcv_PayArray(1))<>"" then
					if IsNumeric(trim(pcv_PayArray(1))) then
						PayFees=cdbl(trim(pcv_PayArray(1)))
						if PayFees>0 then
							tmpS=tmpS & "Fees: " & scCurSign & money(PayFees) & vbcrlf
						end if
					end if
				end if
			else
				tmpS=tmpS & "Payment Details: " & pcv_paymentDetails & vbcrlf
			end if
			if pcv_paymentCode<>"" then
			tmpS=tmpS & "Payment Gateway: " & pcv_paymentCode & vbcrlf
			end if
			if pcv_gwTransId<>"" then
			tmpS=tmpS & "Transaction ID: " & pcv_gwTransId & vbcrlf
			end if
			if pcv_gwAuthCode<>"" then
			tmpS=tmpS & "Authorization Code: " & pcv_gwAuthCode
			end if
			tmpS=tmpS & chr(34) & ","
			StringBuilderObj.append tmpS
End If
if chk_idAffiliate="1" then 
StringBuilderObj.append chr(34)  &pcv_idAffiliate & chr(34) & ","
End If
if chk_AffiliateName="1" then
	StringBuilderObj.append chr(34)
	if pcv_idAffiliate>"1" then
			
			mySQL="select affiliateName,commission FROM affiliates where idAffiliate=" & pcv_idAffiliate
			set rs=connTemp.execute(mySQL)
			
			StringBuilderObj.append rs("affiliateName") & " (#" & pcv_idAffiliate & ")"
			
			set rs=nothing
			
	end if
	StringBuilderObj.append chr(34) & ","
End If
if chk_affiliatePay="1" then
	StringBuilderObj.append chr(34) & scCurSign & money(pcv_affiliatePay) & chr(34) & ","
End If
if chk_iRewardPoints="1" then 
StringBuilderObj.append chr(34)  &pcv_iRewardPoints & chr(34) & ","
End If
if chk_iRewardPointsCustAccrued="1" then 
StringBuilderObj.append chr(34)  &pcv_iRewardPointsCustAccrued & chr(34) & ","
End If
if chk_IDRefer="1" then 
StringBuilderObj.append chr(34)  &pcv_IDRefer & chr(34) & ","
End If
if chk_ReferName="1" then
	if pcv_IDRefer<>"0" then
		query="SELECT name FROM REFERRER where IdRefer="&pcv_IDRefer
		set rsReferObj=Server.CreateObject("ADODB.RecordSet")
		set rsReferObj=conntemp.execute(query)
		if NOT rsReferObj.eof then
			pcv_ReferName=rsReferObj("name")
		else
			pcv_ReferName=""
		end if
		set rsReferObj=nothing
	else
		pcv_ReferName=""
	end if
	StringBuilderObj.append chr(34)  &pcv_ReferName & chr(34) & ","
End If
if chk_rmaCredit="1" then 
StringBuilderObj.append chr(34)  &pcv_rmaCredit & chr(34) & ","
End If
if chk_gwAuthCode="1" then
StringBuilderObj.append chr(34) & pcv_gwAuthCode & chr(34) & ","
End If
if chk_gwTransId="1" then
StringBuilderObj.append chr(34) & pcv_gwTransId & chr(34) & ","
End If
if chk_paymentCode="1" then
StringBuilderObj.append chr(34) & pcv_paymentCode & chr(34) & ","
End If
if chk_taxAmount="1" then 
StringBuilderObj.append chr(34)  &scCurSign & money(pcv_taxAmount) & chr(34) & ","
End If
if chk_taxDetails="1" then
tmpS=""
tmpS=tmpS & chr(34) & "Tax Amount: " & scCurSign & money(pcv_taxAmount) & vbcrlf
			IF cdbl(pcv_taxAmount)>0 then
				if instr(pcv_taxDetails,",")>0 then
				TaxArray=split(pcv_taxDetails,",")
				For m=0 to (ubound(TaxArray)-1)
					tmpTax=split(TaxArray(m),"|")
					tmpS=tmpS & ucase(tmpTax(0)) & " - Amount: " & scCurSign & money(tmpTax(1)) & vbcrlf
				Next
				end if
			END IF
tmpS=tmpS & chr(34) & ","
StringBuilderObj.append tmpS
End If
if chk_ord_VAT="1" then 
StringBuilderObj.append chr(34)  &scCurSign & money(pcv_ord_VAT) & chr(34) & ","
End If
if chk_pcOrd_DiscountDetails="1" then
tmpS=""
tmpS=tmpS & chr(34)
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
				tmpS=tmpS & "Discount Name: " & replace(trim(pcv_DisArray(0)),chr(34),chr(34) & chr(34)) & vbcrlf
				if DisAmount<>0 then
					tmpS=tmpS & "Amount: -" & scCurSign & money(DisAmount) & vbcrlf 
				end if
			end if
			Next
			
			end if
			
			if pcv_pcOrd_CatDiscounts<>"" then
			if cdbl(pcv_pcOrd_CatDiscounts)<>0 then
			tmpS=tmpS & "Discount by Categories: -" & scCurSign & money(pcv_pcOrd_CatDiscounts)
			end if
			end if
tmpS=tmpS & chr(34) & ","
StringBuilderObj.append tmpS
End If
if chk_pcOrd_CatDiscounts="1" then 
StringBuilderObj.append chr(34)  &"-" & scCurSign & money(pcv_pcOrd_CatDiscounts) & chr(34) & ","
End If
if chk_pcOrd_GiftCertificates="1" then
	
	tmpS=""
	tmpS=tmpS & chr(34)
    if pcv_pcOrd_RedeemedGC<>"0" then
		GCArry=split(pcv_pcOrd_RedeemedGC,"|g|")
		intArryCnt=ubound(GCArry)
	
		for k=0 to intArryCnt
	
		if GCArry(k)<>"" then
			GCInfo = split(GCArry(k),"|s|")
			if GCInfo(2)="" OR IsNull(GCInfo(2)) then
				GCInfo(2)=0
			end if
			pGiftCode=GCInfo(0)
			pGiftUsed=GCInfo(2)
		query="SELECT products.IDProduct,products.Description FROM pcGCOrdered,Products WHERE products.idproduct=pcGCOrdered.pcGO_idproduct AND pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
		SET rsGC=server.CreateObject("ADODB.RecordSet")
		SET rsGC=connTemp.execute(query)
	
		if NOT rsGC.eof then
			pIdproduct=rsGC("idproduct")
			pName=rsGC("Description")
			pCode=pGiftCode
			tmpS=tmpS & "GC Name: " & pName & vbcrlf
				query="SELECT pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status FROM pcGCOrdered WHERE pcGO_GcCode='" & pGiftCode & "'"
				SET rsGCObj=server.CreateObject("ADODB.RecordSet")
				SET rsGCObj=connTemp.execute(query)
				
				if NOT rsGCObj.eof then
					pcGO_GcCode=rsGCObj("pcGO_GcCode")
					pExpDate=rsGCObj("pcGO_ExpDate")
					pGCAmount=rsGCObj("pcGO_Amount")
					pGCStatus=rsGCObj("pcGO_Status")
	
					tmpS=tmpS & "GC Code: " & pcGO_GcCode & vbcrlf
					tmpS=tmpS & "Used for this order: " & scCurSign & money(pGiftUsed) & vbcrlf
					if cdbl(pGCAmount)<=0 then
						tmpS=tmpS & "Completely Redeemed" & vbcrlf
					else
						tmpS=tmpS & "Available Amount: " & scCurSign & money(pGCAmount) & vbcrlf
						if year(pExpDate)="1900" then
							tmpS=tmpS & "GC does not expire" & vbcrlf
						else
							if scDateFrmt="DD/MM/YY" then
								pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
							else
								pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
							end if
							tmpS=tmpS & "Expiring: " & pExpDate & vbcrlf
						end if
						if pGCStatus="1" then
							tmpS=tmpS & "Status: Active" & vbcrlf
						else
							tmpS=tmpS & "Status: Inactive" & vbcrlf
						end if
					end if
				end if
				set rsGCObj=nothing
	
			end if
			set rsGC=nothing
		end if
		Next
	end if
	tmpS=tmpS & chr(34) & ","
	StringBuilderObj.append tmpS
end if
if chk_comments="1" then  
StringBuilderObj.append chr(34)  &pcv_comments & chr(34) & ","
End If
if chk_adminComments="1" then  
StringBuilderObj.append chr(34) & pcv_admincomments & chr(34) & ","
End If
if chk_returnDate="1" then
StringBuilderObj.append chr(34) & pcv_returnDate & chr(34) & ","
End If
if chk_returnReason="1" then  
StringBuilderObj.append chr(34)  & pcv_returnReason & chr(34) & ","
End If
if chk_DSNotify="1" then
StringBuilderObj.append chr(34)  & GetDSNotifyData(pcv_idOrder,0) & chr(34) & ","
End if
			a.Writeline(StringBuilderObj.toString())
			set StringBuilderObj = nothing
Next 'Loop Records
	set rstemp=nothing
	End If
	a.Close
	Set fs=Nothing
	closedb()
	response.redirect "getFile.asp?file="& strFile &"&Type=csv"	
End Function

Function CreateXlsFile()
	Dim xlWorkSheet
	Dim xlApplication 
	Set xlApplication=CreateObject("Excel.Application")
	xlApplication.Visible=False
	xlApplication.Workbooks.Add
	Set xlWorksheet=xlApplication.Worksheets(1)
	t=0
if chk_idOrder="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Order ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ord_OrderName="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Order Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_orderDate="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Order Date"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_idCustomer="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Customer ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_CustomerDetails="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Customer Details"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_details="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Order Details"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_total="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Order Total"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_processDate="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Processed Date"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ShippingFullName="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingCompany="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Company"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingAddress="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Address"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingAddress2="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Address 2"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingCity="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping City"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingStateCode="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping State"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingState="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Province"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingCountryCode="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Country"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingZip="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Zip"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shippingPhone="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Phone"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ShipmentDetails="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipment Details"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ordShiptype="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Type"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ordPackageNum="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Number of packages"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shipDate="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipping Date"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_shipVia="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Shipped Via"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_trackingNum="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Tracking Number"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ord_DeliveryDate="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Delivery Date"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_orderStatus="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Order Status"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_PaymentDetails="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Payment Details"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_idAffiliate="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Affiliate ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_AffiliateName="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Affiliate Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_affiliatePay="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Affiliate Payment"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_iRewardPoints="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value=RewardsLabel
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_iRewardPointsCustAccrued="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Accrued " & RewardsLabel
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_IDRefer="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Referrer ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ReferName="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Referrer Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_rmaCredit="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="RMA Credit"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_gwAuthCode="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Authorization Code"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_gwTransId="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Transaction ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_paymentCode="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Payment Gateway"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_taxAmount="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Tax Amount"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_taxDetails="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Tax Details"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_ord_VAT="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="VAT"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_pcOrd_DiscountDetails="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Discount Details"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_pcOrd_CatDiscounts="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Categories Discounts"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_pcOrd_GiftCertificates="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Redeemed Gift Certificates"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_comments="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Customer Comments"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_adminComments="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Admin Comments"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_returnDate="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Return Date"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_returnReason="1" then 
t=t+1
		xlWorksheet.Cells(1,t).Value="Return Reason"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
if chk_DSNotify="1" then
t=t+1
		xlWorksheet.Cells(1,t).Value="Drop-shipper Notifications"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
End If
	iRow=2
if not rstemp.eof then		 
	pcArr=rstemp.getRows()
	set rstemp=nothing
	intCount=ubound(pcArr,2)
	For nk=0 to intCount
	pcv_idOrder=pcArr(0,nk)
	pcv_ShowID=scpre+int(pcv_idOrder)
	pcv_orderDate=pcArr(1,nk)
	pcv_idCustomer=pcArr(2,nk)
	pcv_details=pcArr(58,nk)
	pcv_total=pcArr(4,nk)
	if pcv_total<>"" then
	else
		pcv_total="0"
	end if
	pcv_address=pcArr(5,nk)
	pcv_zip=pcArr(6,nk)
	pcv_stateCode=pcArr(7,nk)
	pcv_state=pcArr(8,nk)
	pcv_city=pcArr(9,nk)
	pcv_countryCode=pcArr(10,nk)
	pcv_comments=pcArr(59,nk)
	pcv_taxAmount=pcArr(12,nk)
	if pcv_taxAmount<>"" then
	else
		pcv_taxAmount="0"
	end if
	pcv_shipmentDetails=pcArr(13,nk)
	pcv_paymentDetails=pcArr(14,nk)
	pcv_discountDetails=pcArr(15,nk)
	pcv_randomNumber=pcArr(16,nk)
	pcv_shippingAddress=pcArr(17,nk)
	pcv_shippingStateCode=pcArr(18,nk)
	pcv_shippingState=pcArr(19,nk)
	pcv_shippingCity=pcArr(20,nk)
	pcv_shippingCountryCode=pcArr(21,nk)
	pcv_shippingZip=pcArr(22,nk)
	pcv_orderStatus=pcArr(23,nk)
	pcv_shippingPhone=pcArr(24,nk)
	pcv_idAffiliate=pcArr(25,nk)
	pcv_processDate=pcArr(26,nk)

	if (chk_shipDate="1") OR (chk_ShipmentDetails="1") OR (chk_shipVia="1") OR (chk_trackingNum="1") then

			query="SELECT pcPackageInfo_ID,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber FROM pcPackageInfo WHERE idorder=" & pcv_idOrder
			set rsStr=connTemp.execute(query)
			pcv_shipDate=""
			pcv_shipVia=""
			pcv_trackingNum=""
			tmp_HavePacks=0
			if not rsStr.eof then
			do while not rsStr.eof
				tmp_HavePacks=1
				tmp_packID=rsStr("pcPackageInfo_ID")
				tmp_packMethod=rsStr("pcPackageInfo_ShipMethod")
				tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
				if scDateFrmt="DD/MM/YY" then
					tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
				end if
				tmp_TrackingNumber=rsStr("pcPackageInfo_TrackingNumber")
				
				if  pcv_shipDate<>"" then
					pcv_shipDate=pcv_shipDate & vbcrlf
				end if
				pcv_shipDate=pcv_shipDate & tmp_processDate
				
				if  pcv_shipVia<>"" then
					pcv_shipVia=pcv_shipVia & vbcrlf
				end if
				pcv_shipVia=pcv_shipVia & "Package ID# " & tmp_packID & " ** " & tmp_packMethod & " ** " & tmp_processDate & " ** " & tmp_TrackingNumber
				
				if  pcv_trackingNum<>"" then
					pcv_trackingNum=pcv_trackingNum & vbcrlf
				end if
				pcv_trackingNum=pcv_trackingNum & tmp_TrackingNumber
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
				pcv_shipVia=pcArr(28,nk)
				pcv_trackingNum=pcArr(29,nk)
			end if
			set rsStr=nothing
			
	end if
			
	pcv_affiliatePay=pcArr(30,nk)
	pcv_returnDate=pcArr(31,nk)
	pcv_returnReason=pcArr(32,nk)
	pcv_iRewardPoints=pcArr(33,nk)
	if pcv_iRewardPoints<>"" then
	else
		pcv_iRewardPoints="0"
	end if
	pcv_ShippingFullName=pcArr(34,nk)
	pcv_iRewardValue=pcArr(35,nk)
	pcv_iRewardRefId=pcArr(36,nk)
	pcv_iRewardPointsRef=pcArr(37,nk)
	pcv_iRewardPointsCustAccrued=pcArr(38,nk)
	pcv_IDRefer=pcArr(39,nk)
	if isNull(pcv_IDRefer) or pcv_IDRefer="" then
		pcv_IDRefer="0"
	end if
	pcv_address2=pcArr(40,nk)
	pcv_shippingCompany=pcArr(41,nk)
	pcv_shippingAddress2=pcArr(42,nk)
	pcv_taxDetails=pcArr(61,nk)
	pcv_adminComments=pcArr(60,nk)
	pcv_rmaCredit=pcArr(45,nk)
	if pcv_rmaCredit<>"" then
	else
		pcv_rmaCredit="0"
	end if
	pcv_DPs=pcArr(46,nk)
	pcv_gwAuthCode=pcArr(47,nk)
	pcv_gwTransId=pcArr(48,nk)
	pcv_paymentCode=pcArr(49,nk)
	pcv_SRF=pcArr(50,nk)
	pcv_ordShiptype=pcArr(51,nk)
	if pcv_ordShiptype<>"" then
	else
		pcv_ordShiptype="0"
	end if
	pcv_ordPackageNum=pcArr(52,nk)
	if pcv_ordPackageNum<>"" then
	else
		pcv_ordPackageNum="0"
	end if
	pcv_ord_DeliveryDate=pcArr(53,nk)
	pcv_ord_OrderName=pcArr(54,nk)
	pcv_ord_VAT=pcArr(55,nk)
	if pcv_ord_VAT<>"" then
	else
		pcv_ord_VAT="0"
	end if

	pcv_pcOrd_CatDiscounts=pcArr(56,nk)
	if pcv_pcOrd_CatDiscounts<>"" then
	else
		pcv_pcOrd_CatDiscounts="0"
	end if
	
	pcv_pcOrd_RedeemedGC=pcArr(57,nk)
	if pcv_pcOrd_RedeemedGC<>"" then
	else
		pcv_pcOrd_RedeemedGC="0"
	end if
	
	t=0
	if chk_idOrder="1" then 
		t=t+1
		xlWorksheet.Cells(iRow,t).Value=pcv_ShowID 
	End If
if chk_ord_OrderName="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_ord_OrderName 
End If
if chk_orderDate="1" then
dtOrderDate=pcv_orderDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if
	t=t+1
				xlWorksheet.Cells(iRow,t).Value= dtOrderDate 
End If
if chk_idCustomer="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_IDCustomer 
End If
if chk_CustomerDetails="1" then
	mySQL="select name,lastname,customerCompany,phone,email FROM customers where idcustomer=" & pcv_idCustomer
	set rs=connTemp.execute(mySQL)
			
	pcv_CustName=rs("name") & " " & rs("lastname") & " (#" & pcv_idcustomer & ")"
	pcv_CustCompany=rs("customerCompany")
	pcv_CustPhone=rs("phone")
	pcv_CustEmail=rs("email")
	
	set rs=nothing 
	
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
	if pcv_CustEmail<>"" then
		tmpS1=tmpS1 & vbcrlf & "Email: " & pcv_CustEmail
	end if
t=t+1
				xlWorksheet.Cells(iRow,t).Value=tmpS1
End If
if chk_details="1" then 
	OrdDetails=replace(GetProductCFs(pcv_idOrder,pcv_details)," ||",""&scCurSign&"")
t=t+1
				xlWorksheet.Cells(iRow,t).Value=OrdDetails
End If
if chk_total="1" then
t=t+1
				xlWorksheet.Cells(iRow,t).Value= scCurSign & money(pcv_total) 
End If
if chk_processDate="1" then
dtOrderDate=pcv_processDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if
t=t+1
				xlWorksheet.Cells(iRow,t).Value=dtOrderDate
End If
if chk_ShippingFullName="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_ShippingFullName 
End If
if chk_shippingCompany="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingCompany 
End If
if chk_shippingAddress="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingAddress 
End If
if chk_shippingAddress2="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingAddress2 
End If
if chk_shippingCity="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingCity 
End If
if chk_shippingStateCode="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingStateCode 
End If
if chk_shippingState="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingState 
End If
if chk_shippingCountryCode="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingCountryCode 
End If
if chk_shippingZip="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingZip 
End If
if chk_shippingPhone="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_shippingPhone 
End If
if chk_ShipmentDetails="1" then 
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
			
			tmpS=""
			
			tmpS=tmpS & "Shipping Method: " & ShipDetails & vbcrlf
			if pcv_ordShiptype="0" then
			tmpS=tmpS & "Shipping Type: Residential" & vbcrlf
			end if
			if pcv_ordShiptype="1" then
			tmpS=tmpS & "Shipping Type: Commercial" & vbcrlf
			end if
			if ShipFees>"0" then
			tmpS=tmpS & "Fees: " & scCurSign & money(ShipFees) & vbcrlf
			end if
			if HandlingFees>"0" then
			tmpS=tmpS & "Handling Fees: " & scCurSign & money(HandlingFees) & vbcrlf
			end if
			if pcv_ShipVia<>"" then
				if tmp_HavePacks=0 then
					tmpS=tmpS & "Shipped Via: " & pcv_ShipVia & vbcrlf
				end if
			end if
			if pcv_ordPackageNum<>"" then
			tmpS=tmpS & "Number of packages: " & pcv_ordPackageNum & vbcrlf
			end if
			if pcv_shipDate<>"" then
				dtOrderDate=pcv_shipDate
				tmpS=tmpS & "Date Shipped: " & dtOrderDate & vbcrlf
			end if
			if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			tmpS=tmpS & "Delivery Date: " & dtOrderDate
			end if
			
			t=t+1
				xlWorksheet.Cells(iRow,t).Value=tmpS
End If
if chk_ordShiptype="1" then
			tmpS=""
			if pcv_ordShiptype="0" then
			tmpS=tmpS & "Shipping Type: Residential"
			end if
			if pcv_ordShiptype="1" then
			tmpS=tmpS & "Shipping Type: Commercial"
			end if
			t=t+1
				xlWorksheet.Cells(iRow,t).Value=tmpS
End If
if chk_ordPackageNum="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_ordPackageNum 
End If
if chk_shipDate="1" then
	strtext=""
	if pcv_shipDate<>"" then
			dtOrderDate=pcv_shipDate
			strtext=strtext & dtOrderDate
	end if
	t=t+1
				xlWorksheet.Cells(iRow,t).Value=strtext
End If
if chk_shipVia="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_ShipVia 
End If
if chk_trackingNum="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_trackingNum 
End If
if chk_ord_DeliveryDate="1" then
	strtext=""
	if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			strtext=strtext & dtOrderDate
	end if
	t=t+1
				xlWorksheet.Cells(iRow,t).Value=strtext
End If
if chk_orderStatus="1" then
					OrderStatusStr="Incomplete"		
					Select case cint(pcv_orderStatus)
					Case 2: OrderStatusStr="Pending"
					Case 3: OrderStatusStr="Processed"
					Case 4: OrderStatusStr="Shipped"
					Case 5: OrderStatusStr="Canceled"
					Case 6: OrderStatusStr="Return"
					Case 7: OrderStatusStr="Partially Shipped"
					Case 8: OrderStatusStr="Shipping"
					Case 10: OrderStatusStr="Delivered"
					Case 12: OrderStatusStr="Archived"
					End Select
				t=t+1
				xlWorksheet.Cells(iRow,t).Value=OrderStatusStr
End If
if chk_PaymentDetails="1" then
			tmpS=""
			if instr(pcv_paymentDetails,"||")>0 then
				pcv_PayArray=split(pcv_paymentDetails,"||")
				tmpS=tmpS & "Payment Method: " & trim(pcv_PayArray(0)) & vbcrlf
				if trim(pcv_PayArray(1))<>"" then
					if IsNumeric(trim(pcv_PayArray(1))) then
						PayFees=cdbl(trim(pcv_PayArray(1)))
						if PayFees>0 then
							tmpS=tmpS & "Fees: " & scCurSign & money(PayFees) & vbcrlf
						end if
					end if
				end if
			else
				tmpS=tmpS & "Payment Details: " & pcv_paymentDetails & vbcrlf
			end if
			if pcv_paymentCode<>"" then
			tmpS=tmpS & "Payment Gateway: " & pcv_paymentCode & vbcrlf
			end if
			if pcv_gwTransId<>"" then
			tmpS=tmpS & "Transaction ID: " & pcv_gwTransId & vbcrlf
			end if
			if pcv_gwAuthCode<>"" then
			tmpS=tmpS & "Authorization Code: " & pcv_gwAuthCode
			end if
			t=t+1
				xlWorksheet.Cells(iRow,t).Value=tmpS
End If
if chk_idAffiliate="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_idAffiliate 
End If
if chk_AffiliateName="1" then
	strtext=""
	if pcv_idAffiliate>"1" then
			
			mySQL="select affiliateName,commission FROM affiliates where idAffiliate=" & pcv_idAffiliate
			set rs=connTemp.execute(mySQL)
			
			strtext=strtext & rs("affiliateName") & " (#" & pcv_idAffiliate & ")"
			
			set rs=nothing
			
	end if
	t=t+1
				xlWorksheet.Cells(iRow,t).Value=strtext
End If
if chk_affiliatePay="1" then
	t=t+1
				xlWorksheet.Cells(iRow,t).Value=scCurSign & money(pcv_affiliatePay) 
End If
if chk_iRewardPoints="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_iRewardPoints 
End If
if chk_iRewardPointsCustAccrued="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_iRewardPointsCustAccrued 
End If
if chk_IDRefer="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_IDRefer 
End If
if chk_ReferName="1" then
	if pcv_IDRefer<>"0" then
		query="SELECT name FROM REFERRER where IdRefer="&pcv_IDRefer
		set rsReferObj=Server.CreateObject("ADODB.RecordSet")
		set rsReferObj=conntemp.execute(query)
		if NOT rsReferObj.eof then
			pcv_ReferName=rsReferObj("name")
		else
			pcv_ReferName=""
		end if
		set rsReferObj=nothing
	else
		pcv_ReferName=""
	end if
	t=t+1
	xlWorksheet.Cells(iRow,t).Value=pcv_ReferName
End If
if chk_rmaCredit="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_rmaCredit 
End If
if chk_gwAuthCode="1" then
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_gwAuthCode 
End If
if chk_gwTransId="1" then
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_gwTransId 
End If
if chk_paymentCode="1" then
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_paymentCode 
End If
if chk_taxAmount="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=scCurSign & money(pcv_taxAmount) 
End If
if chk_taxDetails="1" then
tmpS=""
tmpS=tmpS & "Tax Amount: " & scCurSign & money(pcv_taxAmount) & vbcrlf
			IF cdbl(pcv_taxAmount)>0 then
				if instr(pcv_taxDetails,",")>0 then
				TaxArray=split(pcv_taxDetails,",")
				For m=0 to (ubound(TaxArray)-1)
					tmpTax=split(TaxArray(m),"|")
					tmpS=tmpS & ucase(tmpTax(0)) & " - Amount: " & scCurSign & money(tmpTax(1)) & vbcrlf
				Next
				end if
			END IF
t=t+1
				xlWorksheet.Cells(iRow,t).Value=tmpS
End If
if chk_ord_VAT="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value=scCurSign & money(pcv_ord_VAT) 
End If
if chk_pcOrd_DiscountDetails="1" then
tmpS=""
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
				tmpS=tmpS & "Discount Name: " & replace(trim(pcv_DisArray(0)),chr(34),chr(34) & chr(34)) & vbcrlf
				if DisAmount<>0 then
					tmpS=tmpS & "Amount: -" & scCurSign & money(DisAmount) & vbcrlf 
				end if
			end if
			Next
			
			end if
			
			if pcv_pcOrd_CatDiscounts<>"" then
			if cdbl(pcv_pcOrd_CatDiscounts)<>0 then
			tmpS=tmpS & "Discount by Categories: -" & scCurSign & money(pcv_pcOrd_CatDiscounts)
			end if
			end if
t=t+1
				xlWorksheet.Cells(iRow,t).Value=tmpS
End If
if chk_pcOrd_CatDiscounts="1" then 
t=t+1
				xlWorksheet.Cells(iRow,t).Value="-" & scCurSign & money(pcv_pcOrd_CatDiscounts) 
End If
if chk_pcOrd_GiftCertificates="1" then
tmpS=""
	if pcv_pcOrd_RedeemedGC<>"0" then
		GCArry=split(pcv_pcOrd_RedeemedGC,"|g|")
		intArryCnt=ubound(GCArry)
	
		for k=0 to intArryCnt
	
		if GCArry(k)<>"" then
			GCInfo = split(GCArry(k),"|s|")
			if GCInfo(2)="" OR IsNull(GCInfo(2)) then
				GCInfo(2)=0
			end if
			pGiftCode=GCInfo(0)
			pGiftUsed=GCInfo(2)
		query="SELECT products.IDProduct,products.Description FROM pcGCOrdered,Products WHERE products.idproduct=pcGCOrdered.pcGO_idproduct AND pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
		SET rsGC=server.CreateObject("ADODB.RecordSet")
		SET rsGC=connTemp.execute(query)
	
		if NOT rsGC.eof then
			pIdproduct=rsGC("idproduct")
			pName=rsGC("Description")
			pCode=pGiftCode
			tmpS=tmpS & "GC Name: " & pName & vbcrlf
				query="SELECT pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status FROM pcGCOrdered WHERE pcGO_GcCode='" & pGiftCode & "'"
				SET rsGCObj=server.CreateObject("ADODB.RecordSet")
				SET rsGCObj=connTemp.execute(query)
				
				if NOT rsGCObj.eof then
					pcGO_GcCode=rsGCObj("pcGO_GcCode")
					pExpDate=rsGCObj("pcGO_ExpDate")
					pGCAmount=rsGCObj("pcGO_Amount")
					pGCStatus=rsGCObj("pcGO_Status")
	
					tmpS=tmpS & "GC Code: " & pcGO_GcCode & vbcrlf
					tmpS=tmpS & "Used for this order: " & scCurSign & money(pGiftUsed) & vbcrlf
					if cdbl(pGCAmount)<=0 then
						tmpS=tmpS & "Completely Redeemed" & vbcrlf
					else
						tmpS=tmpS & "Available Amount: " & scCurSign & money(pGCAmount) & vbcrlf
						if year(pExpDate)="1900" then
							tmpS=tmpS & "GC does not expire" & vbcrlf
						else
							if scDateFrmt="DD/MM/YY" then
								pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
							else
								pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
							end if
							tmpS=tmpS & "Expiring: " & pExpDate & vbcrlf
						end if
						if pGCStatus="1" then
							tmpS=tmpS & "Status: Active" & vbcrlf
						else
							tmpS=tmpS & "Status: Inactive" & vbcrlf
						end if
					end if
				end if
				set rsGCObj=nothing
	
			end if
			set rsGC=nothing
		end if
		Next
	end if
t=t+1
				xlWorksheet.Cells(iRow,t).Value=tmpS
End If
if chk_comments="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_comments 
End If
if chk_adminComments="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_admincomments 
End If
if chk_returnDate="1" then
t=t+1
				xlWorksheet.Cells(iRow,t).Value=pcv_returnDate 
End If
if chk_returnReason="1" then  
t=t+1
				xlWorksheet.Cells(iRow,t).Value= pcv_returnReason 
End If
if chk_DSNotify="1" then
t=t+1
				xlWorksheet.Cells(iRow,t).Value= GetDSNotifyData(pcv_idOrder,0)
End if
			iRow=iRow + 1
			
	Next
End If 'Have records
set rstemp=nothing

	strFile=GenFileName()
	xlWorksheet.SaveAs Server.MapPath(".") & "\" & strFile & ".xls"
	xlApplication.Quit												' Close the Workbook
	Set xlWorksheet=Nothing
	Set xlApplication=Nothing
	response.redirect "getFile.asp?file="& strFile &"&Type=xls"
End Function
%>