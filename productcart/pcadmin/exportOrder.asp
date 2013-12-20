<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Custom Data Export" %>
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
<% 
response.Buffer=true
Response.Expires=0

dim query, queryQ, conntemp, rstemp

call openDb()

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
						tmpStr2=tmpStr2 & "||"
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
chk_shipmentDetails= request.form("chk_shipmentDetails")
chk_AffiliateName= request.form("chk_AffiliateName")
chk_DSNotify=request.form("chk_DSNotify")

err.clear
Dim strTDateVar, strTDateVar2, DateVar, DateVar2
strTDateVar=Request("FromDate")
DateVar=strTDateVar
if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
end if
strTDateVar2=Request("ToDate")
DateVar2=strTDateVar2
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	if err.number<>0 then
		DateVar=Request("FromDate")
		DateVar2=Request("ToDate")
	end if
end if
err.clear

if (DateVar<>"") and IsDate(DateVar) then
    if SQL_Format = "1" then DateVar = day(DateVar) & "/" & month(DateVar) & "/" & year(DateVar)
	if scDB="Access" then
		TempSQL1=" AND orders.orderDate >=#" & DateVar & "# "
	else
		TempSQL1=" AND orders.orderDate >='" & DateVar & "' "
	end if
else
	TempSQL1=""
end if
if (DateVar2<>"") and IsDate(DateVar2) then
    if SQL_Format = "1" then DateVar2 = day(DateVar2) & "/" & month(DateVar2) & "/" & year(DateVar2)
	if scDB="Access" then
		TempSQL2=" AND orders.orderDate <=#" & DateVar2 & "# "
	else
		TempSQL2=" AND orders.orderDate <='" & DateVar2 & "' "
	end if
else
	TempSQL2=""	
end if

Dim intIncludeAll, TempSQLall
intIncludeAll = Request.Form("includeAll")
	if intIncludeAll = "1" then
		TempSQLall ="WHERE (orders.orderStatus>0 AND orders.orderStatus<7) "
		else
		TempSQLall ="WHERE ((orders.orderStatus>1 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) "
	end if

if (DateVar<>"") then 
	query="SELECT idOrder,orderDate,idCustomer,details,total,address,zip,stateCode,state,city,countryCode,comments,taxAmount,shipmentDetails,paymentDetails,discountDetails,randomNumber,shippingAddress,shippingStateCode,shippingState,shippingCity,shippingCountryCode,shippingZip,orderStatus,pcOrd_shippingPhone,idAffiliate,processDate,shipDate,shipVia,trackingNum,affiliatePay,returnDate,returnReason,iRewardPoints,ShippingFullName,iRewardValue,iRewardRefId,iRewardPointsRef,iRewardPointsCustAccrued,IDRefer,address2,shippingCompany,shippingAddress2,taxDetails,adminComments,rmaCredit,DPs,gwAuthCode,gwTransId,paymentCode,SRF,ordShiptype,ordPackageNum,ord_DeliveryDate,ord_OrderName,ord_VAT,pcOrd_CatDiscounts FROM orders " & TempSQLall & TempSQL1 & TempSQL2 & " ORDER BY orders.orderDate ASC;"
else
	query="SELECT idOrder,orderDate,idCustomer,details,total,address,zip,stateCode,state,city,countryCode,comments,taxAmount,shipmentDetails,paymentDetails,discountDetails,randomNumber,shippingAddress,shippingStateCode,shippingState,shippingCity,shippingCountryCode,shippingZip,orderStatus,pcOrd_shippingPhone,idAffiliate,processDate,shipDate,shipVia,trackingNum,affiliatePay,returnDate,returnReason,iRewardPoints,ShippingFullName,iRewardValue,iRewardRefId,iRewardPointsRef,iRewardPointsCustAccrued,IDRefer,address2,shippingCompany,shippingAddress2,taxDetails,adminComments,rmaCredit,DPs,gwAuthCode,gwTransId,paymentCode,SRF,ordShiptype,ordPackageNum,ord_DeliveryDate,ord_OrderName,ord_VAT,pcOrd_CatDiscounts FROM orders " & TempSQLall & " ORDER BY orders.orderDate ASC;"
end if

set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(query)

IF rstemp.eof then
set rstemp=nothing
call closedb()
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessage">
			Your search did not return any results.
		</div>
		<p>&nbsp;</p>
		<p>
			<input type=button value=" Back " onclick="javascript:history.back()" class="ibtnGrey">
		</p>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->
<%
response.end
ELSE
HTMLResult=""
set StringBuilderObj = new StringBuilder
if chk_idOrder="1" then 
StringBuilderObj.append "<td><b>" & "Order ID" & "</b></td>"
End If
if chk_ord_OrderName="1" then 
StringBuilderObj.append "<td><b>" & "Order Name" & "</b></td>"
End If
if chk_orderDate="1" then 
StringBuilderObj.append "<td><b>" & "Order Date" & "</b></td>"
End If
if chk_idCustomer="1" then 
StringBuilderObj.append "<td><b>" & "Customer ID" & "</b></td>"
End If
if chk_CustomerDetails="1" then 
StringBuilderObj.append "<td><b>" & "Customer Details" & "</b></td>"
End If
if chk_details="1" then 
StringBuilderObj.append "<td><b>" & "Order Details" & "</b></td>"
End If
if chk_total="1" then 
StringBuilderObj.append "<td><b>" & "Order Total" & "</b></td>"
End If
if chk_processDate="1" then 
StringBuilderObj.append "<td><b>" & "Processed Date" & "</b></td>"
End If
if chk_ShippingFullName="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Name" & "</b></td>"
End If
if chk_shippingCompany="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Company" & "</b></td>"
End If
if chk_shippingAddress="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Address" & "</b></td>"
End If
if chk_shippingAddress2="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Address 2" & "</b></td>"
End If
if chk_shippingCity="1" then 
StringBuilderObj.append "<td><b>" & "Shipping City" & "</b></td>"
End If
if chk_shippingStateCode="1" then 
StringBuilderObj.append "<td><b>" & "Shipping State" & "</b></td>"
End If
if chk_shippingState="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Province" & "</b></td>"
End If
if chk_shippingCountryCode="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Country" & "</b></td>"
End If
if chk_shippingZip="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Zip" & "</b></td>"
End If
if chk_shippingPhone="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Phone" & "</b></td>"
End If
if chk_ShipmentDetails="1" then 
StringBuilderObj.append "<td><b>" & "Shipment Details" & "</b></td>"
End If
if chk_ordShiptype="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Type" & "</b></td>"
End If
if chk_ordPackageNum="1" then 
StringBuilderObj.append "<td><b>" & "Number of packages" & "</b></td>"
End If
if chk_shipDate="1" then 
StringBuilderObj.append "<td><b>" & "Shipping Date" & "</b></td>"
End If
if chk_shipVia="1" then 
StringBuilderObj.append "<td><b>" & "Shipped Via" & "</b></td>"
End If
if chk_trackingNum="1" then 
StringBuilderObj.append "<td><b>" & "Tracking Number" & "</b></td>"
End If
if chk_ord_DeliveryDate="1" then 
StringBuilderObj.append "<td><b>" & "Delivery Date" & "</b></td>"
End If
if chk_orderStatus="1" then 
StringBuilderObj.append "<td><b>" & "Order Status" & "</b></td>"
End If
if chk_PaymentDetails="1" then 
StringBuilderObj.append "<td><b>" & "Payment Details" & "</b></td>"
End If
if chk_idAffiliate="1" then 
StringBuilderObj.append "<td><b>" & "Affiliate ID" & "</b></td>"
End If
if chk_AffiliateName="1" then 
StringBuilderObj.append "<td><b>" & "Affiliate Name" & "</b></td>"
End If
if chk_affiliatePay="1" then 
StringBuilderObj.append "<td><b>" & "Affiliate Payment" & "</b></td>"
End If
if chk_iRewardPoints="1" then 
StringBuilderObj.append "<td><b>" & RewardsLabel & "</b></td>"
End If
if chk_iRewardPointsCustAccrued="1" then 
StringBuilderObj.append "<td><b>" & "Accrued " & RewardsLabel & "</b></td>"
End If
if chk_IDRefer="1" then 
StringBuilderObj.append "<td><b>" & "Referrer ID" & "</b></td>"
End If
if chk_ReferName="1" then 
StringBuilderObj.append "<td><b>" & "Referrer Name" & "</b></td>"
End If
if chk_rmaCredit="1" then 
StringBuilderObj.append "<td><b>" & "RMA Credit" & "</b></td>"
End If
if chk_gwAuthCode="1" then 
StringBuilderObj.append "<td><b>" & "Authorization Code" & "</b></td>"
End If
if chk_gwTransId="1" then 
StringBuilderObj.append "<td><b>" & "Transaction ID" & "</b></td>"
End If
if chk_paymentCode="1" then 
StringBuilderObj.append "<td><b>" & "Payment Gateway" & "</b></td>"
End If
if chk_taxAmount="1" then 
StringBuilderObj.append "<td><b>" & "Tax Amount" & "</b></td>"
End If
if chk_taxDetails="1" then 
StringBuilderObj.append "<td><b>" & "Tax Details" & "</b></td>"
End If
if chk_ord_VAT="1" then 
StringBuilderObj.append "<td><b>" & "VAT" & "</b></td>"
End If
if chk_pcOrd_DiscountDetails="1" then 
StringBuilderObj.append "<td><b>" & "Discount Details" & "</b></td>"
End If
if chk_pcOrd_CatDiscounts="1" then 
StringBuilderObj.append "<td><b>" & "Categories Discounts" & "</b></td>"
End If
if chk_comments="1" then 
StringBuilderObj.append "<td><b>" & "Customer Comments" & "</b></td>"
End If
if chk_adminComments="1" then 
StringBuilderObj.append "<td><b>" & "Admin Comments" & "</b></td>"
End If
if chk_returnDate="1" then 
StringBuilderObj.append "<td><b>" & "Return Date" & "</b></td>"
End If
if chk_returnReason="1" then 
StringBuilderObj.append "<td><b>" & "Return Reason" & "</b></td>"
End If
if chk_DSNotify="1" then 
StringBuilderObj.append "<td><b>" & "Drop-shipper Notifications" & "</b></td>"
End If
HTMLResult="<table><tr>" & StringBuilderObj.toString() & "</tr>"
set StringBuilderObj = nothing

do until rstemp.eof		 
pcv_idOrder=rstemp(0)
pcv_ShowID=scpre+int(pcv_idOrder)
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
					pcv_shipDate=pcv_shipDate & "||"
				end if
				pcv_shipDate=pcv_shipDate & tmp_processDate
				
				if  pcv_shipVia<>"" then
					pcv_shipVia=pcv_shipVia & "||"
				end if
				pcv_shipVia=pcv_shipVia & "Package ID# " & tmp_packID & " ** " & tmp_packMethod & " ** " & tmp_processDate & " ** " & tmp_TrackingNumber
				
				if  pcv_trackingNum<>"" then
					pcv_trackingNum=pcv_trackingNum & "||"
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
				pcv_shipVia=rstemp(28)
				pcv_trackingNum=rstemp(29)
			end if
			set rsStr=nothing
			
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

if chk_idOrder="1" then 
StringBuilderObj.append "<td>"  &pcv_ShowID & "</td>"
End If
if chk_ord_OrderName="1" then  
StringBuilderObj.append "<td>"  &pcv_ord_OrderName & "</td>"
End If
if chk_orderDate="1" then
dtOrderDate=pcv_orderDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if
	StringBuilderObj.append "<td>"  & dtOrderDate & "</td>"
End If
if chk_idCustomer="1" then 
StringBuilderObj.append "<td>"  &pcv_IDCustomer & "</td>"
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
	tmpS1=pcv_CustName & "||"
	if trim(pcv_CustCompany)<>"" then
		tmpS1=tmpS1 & "Company: " & pcv_CustCompany & "||"
	end if
	tmpS1=tmpS1 & pcv_address & "||"
	if trim(pcv_address2)<>"" then
		tmpS1=tmpS1 & "Address 2: " & pcv_address2 & "||"
	end if
	tmpS1=tmpS1 & pcv_city &", " & pcv_statecode & pcv_state & " " & pcv_zip & ", " & pcv_countryCode & "||"
	if pcv_CustPhone<>"" then
		tmpS1=tmpS1 & "Phone: " & pcv_CustPhone
	end if
	if pcv_CustEmail<>"" then
		tmpS1=tmpS1 & "||" & "Email: " & pcv_CustEmail
	end if
	StringBuilderObj.append "<td>" & tmpS1 & "</td>"
End If
if chk_details="1" then 
	OrdDetails=replace(pcv_details," ||",""&scCurSign&"")
	StringBuilderObj.append "<td>" & OrdDetails & "</td>"
End If
if chk_total="1" then
	StringBuilderObj.append "<td>" & scCurSign & money(pcv_total) & "</td>"
End If
if chk_processDate="1" then
dtOrderDate=pcv_processDate
if scDateFrmt="DD/MM/YY" then
	dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
end if
StringBuilderObj.append "<td>" & dtOrderDate & "</td>"
End If
if chk_ShippingFullName="1" then  
StringBuilderObj.append "<td>"  &pcv_ShippingFullName & "</td>"
End If
if chk_shippingCompany="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingCompany & "</td>"
End If
if chk_shippingAddress="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingAddress & "</td>"
End If
if chk_shippingAddress2="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingAddress2 & "</td>"
End If
if chk_shippingCity="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingCity & "</td>"
End If
if chk_shippingStateCode="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingStateCode & "</td>"
End If
if chk_shippingState="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingState & "</td>"
End If
if chk_shippingCountryCode="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingCountryCode & "</td>"
End If
if chk_shippingZip="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingZip & "</td>"
End If
if chk_shippingPhone="1" then  
StringBuilderObj.append "<td>"  &pcv_shippingPhone & "</td>"
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
			
			tmpS=tmpS & "<td>" & "Shipping Method: " & ShipDetails & "||"
			if pcv_ordShiptype="0" then
			tmpS=tmpS & "Shipping Type: Residential" & "||"
			end if
			if pcv_ordShiptype="1" then
			tmpS=tmpS & "Shipping Type: Commercial" & "||"
			end if
			if ShipFees>"0" then
			tmpS=tmpS & "Fees: " & scCurSign & money(ShipFees) & "||"
			end if
			if HandlingFees>"0" then
			tmpS=tmpS & "Handling Fees: " & scCurSign & money(HandlingFees) & "||"
			end if
			if pcv_ShipVia<>"" then
				if tmp_HavePacks=0 then
					tmpS=tmpS & "Shipped Via: " & pcv_ShipVia & "||"
				end if
			end if
			if pcv_ordPackageNum<>"" then
			tmpS=tmpS & "Number of packages: " & pcv_ordPackageNum & "||"
			end if
			if pcv_shipDate<>"" then
				dtOrderDate=pcv_shipDate
				tmpS=tmpS & "Date Shipped: " & dtOrderDate & "||"
			end if
			if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			tmpS=tmpS & "Delivery Date: " & dtOrderDate
			end if
			
			tmpS=tmpS & "</td>"
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
			StringBuilderObj.append "<td>" & tmpS & "</td>"
End If
if chk_ordPackageNum="1" then 
StringBuilderObj.append "<td>" & pcv_ordPackageNum & "</td>"
End If
if chk_shipDate="1" then
	StringBuilderObj.append "<td>"
	if pcv_shipDate<>"" then
			dtOrderDate=pcv_shipDate
			StringBuilderObj.append dtOrderDate
	end if
	StringBuilderObj.append "</td>"
End If
if chk_shipVia="1" then  
StringBuilderObj.append "<td>" & pcv_ShipVia & "</td>"
End If
if chk_trackingNum="1" then 
StringBuilderObj.append "<td>" & pcv_trackingNum & "</td>"
End If
if chk_ord_DeliveryDate="1" then
	StringBuilderObj.append "<td>"
	if pcv_ord_DeliveryDate<>"" then
			dtOrderDate=pcv_ord_DeliveryDate
			if scDateFrmt="DD/MM/YY" then
				dtOrderDate=(day(dtOrderDate)&"/"&month(dtOrderDate)&"/"&year(dtOrderDate))
			end if
			StringBuilderObj.append dtOrderDate
	end if
	StringBuilderObj.append "</td>"
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
					StringBuilderObj.append "<td>" & OrderStatusStr & "</td>"
End If
if chk_PaymentDetails="1" then
			tmpS=""
			tmpS=tmpS & "<td>"
			if instr(pcv_paymentDetails,"||")>0 then
				pcv_PayArray=split(pcv_paymentDetails,"||")
				tmpS=tmpS & "Payment Method: " & trim(pcv_PayArray(0)) & "||"
				if trim(pcv_PayArray(1))<>"" then
					if IsNumeric(trim(pcv_PayArray(1))) then
						PayFees=cdbl(trim(pcv_PayArray(1)))
						if PayFees>0 then
							tmpS=tmpS & "Fees: " & scCurSign & money(PayFees) & "||"
						end if
					end if
				end if
			else
				tmpS=tmpS & "Payment Details: " & pcv_paymentDetails & "||"
			end if
			if pcv_paymentCode<>"" then
			tmpS=tmpS & "Payment Gateway: " & pcv_paymentCode & "||"
			end if
			if pcv_gwTransId<>"" then
			tmpS=tmpS & "Transaction ID: " & pcv_gwTransId & "||"
			end if
			if pcv_gwAuthCode<>"" then
			tmpS=tmpS & "Authorization Code: " & pcv_gwAuthCode
			end if
			tmpS=tmpS & "</td>"
			StringBuilderObj.append tmpS
End If
if chk_idAffiliate="1" then 
StringBuilderObj.append "<td>"  &pcv_idAffiliate & "</td>"
End If
if chk_AffiliateName="1" then
	StringBuilderObj.append "<td>"
	if pcv_idAffiliate>"1" then
			
			mySQL="select affiliateName,commission FROM affiliates where idAffiliate=" & pcv_idAffiliate
			set rs=connTemp.execute(mySQL)
			
			StringBuilderObj.append rs("affiliateName") & " (#" & pcv_idAffiliate & ")"
			
			set rs=nothing
			
	end if
	StringBuilderObj.append "</td>"
End If
if chk_affiliatePay="1" then
	StringBuilderObj.append "<td>" & scCurSign & money(pcv_affiliatePay) & "</td>"
End If
if chk_iRewardPoints="1" then 
StringBuilderObj.append "<td>"  &pcv_iRewardPoints & "</td>"
End If
if chk_iRewardPointsCustAccrued="1" then 
StringBuilderObj.append "<td>"  &pcv_iRewardPointsCustAccrued & "</td>"
End If
if chk_IDRefer="1" then 
StringBuilderObj.append "<td>"  &pcv_IDRefer & "</td>"
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
	StringBuilderObj.append "<td>"  &pcv_ReferName & "</td>"
End If
if chk_rmaCredit="1" then 
StringBuilderObj.append "<td>"  &pcv_rmaCredit & "</td>"
End If
if chk_gwAuthCode="1" then
StringBuilderObj.append "<td>" & pcv_gwAuthCode & "</td>"
End If
if chk_gwTransId="1" then
StringBuilderObj.append "<td>" & pcv_gwTransId & "</td>"
End If
if chk_paymentCode="1" then
StringBuilderObj.append "<td>" & pcv_paymentCode & "</td>"
End If
if chk_taxAmount="1" then 
StringBuilderObj.append "<td>"  &scCurSign & money(pcv_taxAmount) & "</td>"
End If
if chk_taxDetails="1" then
tmpS=""
tmpS=tmpS & "<td>" & "Tax Amount: " & scCurSign & money(pcv_taxAmount) & "||"
			IF cdbl(pcv_taxAmount)>0 then
				if instr(pcv_taxDetails,",")>0 then
				TaxArray=split(pcv_taxDetails,",")
				For m=0 to (ubound(TaxArray)-1)
					tmpTax=split(TaxArray(m),"|")
					tmpS=tmpS & ucase(tmpTax(0)) & " - Amount: " & scCurSign & money(tmpTax(1)) & "||"
				Next
				end if
			END IF
tmpS=tmpS & "</td>"
StringBuilderObj.append tmpS
End If
if chk_ord_VAT="1" then 
StringBuilderObj.append "<td>"  &scCurSign & money(pcv_ord_VAT) & "</td>"
End If
if chk_pcOrd_DiscountDetails="1" then
tmpS=""
tmpS=tmpS & "<td>"
			if instr(pcv_discountDetails,"- ||")>0 then
			
			pcv_DisArray=split(pcv_discountDetails,"- ||")
			DisAmount=cdbl(trim(pcv_DisArray(1)))
			tmpS=tmpS & "Discount Name: " & trim(pcv_DisArray(0)) & "||"
			if DisAmount<>0 then
			tmpS=tmpS & "Amount: -" & scCurSign & money(DisAmount) & "||" 
			end if
			
			end if
			
			if pcv_pcOrd_CatDiscounts<>"" then
			if cdbl(pcv_pcOrd_CatDiscounts)<>0 then
			tmpS=tmpS & "Discount by Categories: -" & scCurSign & money(pcv_pcOrd_CatDiscounts)
			end if
			end if
tmpS=tmpS & "</td>"
StringBuilderObj.append tmpS
End If
if chk_pcOrd_CatDiscounts="1" then 
StringBuilderObj.append "<td>"  &"-" & scCurSign & money(pcv_pcOrd_CatDiscounts) & "</td>"
End If
if chk_comments="1" then  
StringBuilderObj.append "<td>"  &pcv_comments & "</td>"
End If
if chk_adminComments="1" then  
StringBuilderObj.append "<td>" & pcv_admincomments & "</td>"
End If
if chk_returnDate="1" then
StringBuilderObj.append "<td>" & pcv_returnDate & "</td>"
End If
if chk_returnReason="1" then  
StringBuilderObj.append "<td>"  & pcv_returnReason & "</td>"
End If
if chk_DSNotify="1" then  
StringBuilderObj.append "<td>"  & GetDSNotifyData(pcv_idOrder,0) & "</td>"
End If
HTMLResult=HTMLResult & "<tr>" & StringBuilderObj.toString() & "</tr>" & vbcrlf
set StringBuilderObj = nothing
rstemp.moveNext
loop
set rstemp=nothing
HTMLResult=HTMLResult & "</table>"
END IF
closedb()
%>
<% 
Response.ContentType = "application/vnd.ms-excel"
%>
<%=HTMLResult%>