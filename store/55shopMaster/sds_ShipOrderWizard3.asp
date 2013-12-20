<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pageTitle="Shipping Wizard" %>
<% response.Buffer=true %>
<% section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp" --> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="../includes/GoogleCheckout_APIFunctions.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Handler.asp"-->
<!--#include file="AdminHeader.asp"-->
<% Dim connTemp,rs,query
call opendb()

'// Define objects used to create and send Google Checkout Order Processing API requests
Dim xmlRequest
Dim xmlResponse
Dim attrGoogleOrderNumber
Dim elemAmount 
Dim elemReason
Dim elemComment
Dim elemCarrier
Dim elemTrackingNumber
Dim elemMessage
Dim elemSendEmail
Dim elemMerchantOrderNumber
Dim transmitResponse

IF request("action")="add" THEN

	intUSPSLabelFlag=request("USPSLabelOnly") ' value="1"
    pcv_PackageID=request("PackID")
	if pcv_PackageID="" then
		pcv_PackageID=0
	end if
	
	pcv_IdOrder=request("idorder")
	if pcv_IdOrder="" then
		pcv_IdOrder=0
	end if	
	
	pcv_PrdList=request("PrdList")
	
	if (pcv_IdOrder=0) then
		response.redirect "menu.asp"
	end if
	
	query="SELECT orders.pcOrd_GoogleIDOrder FROM orders WHERE idOrder="& pcv_IdOrder
	set rs=server.CreateObject("ADODB.RecordSet")
	Set rs=conntemp.execute(query)
	if Not rs.eof then
		pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder") '// determine if this is a google order
	end if
	set rs=nothing
	
	pcv_method=request("pcv_method")
	pcv_tracking=request("pcv_tracking")
	pcv_shippedDate=request("pcv_shippedDate")
	pcv_AdmComments=request("pcv_AdmComments")
	if pcv_AdmComments<>"" then
		pcv_AdmComments=replace(pcv_AdmComments,"'","''")
	end if
	
	dim dtShippedDate
	dtShippedDate=pcv_shippedDate
	if pcv_shippedDate<>"" then
		if scDateFrmt="DD/MM/YY" then
			err.number=0
			tempShpDt=pcv_shippedDate
			shpDtArr=split(pcv_shippedDate,"/")
			dtShippedDate=(shpDtArr(1)&"/"&shpDtArr(0)&"/"&shpDtArr(2))
			if SQL_Format="1" then
				dtShippedDate=tempShpDt
			end if
		end if
	end if
	
	if intUSPSLabelFlag="1" then
		if scDB="SQL" then
			query="UPDATE pcPackageInfo SET pcPackageInfo_ShipMethod='" & pcv_method & "', pcPackageInfo_ShippedDate='" & dtShippedDate & "', pcPackageInfo_TrackingNumber='" & pcv_tracking & "', pcPackageInfo_Comments='" & pcv_AdmComments & "' WHERE idOrder="&pcv_IdOrder&" AND pcPackageInfo_ID="&pcv_PackageID&";"
		else
			query="UPDATE pcPackageInfo SET pcPackageInfo_ShipMethod='" & pcv_method & "', pcPackageInfo_ShippedDate=#" & dtShippedDate & "#, pcPackageInfo_TrackingNumber='" & pcv_tracking & "', pcPackageInfo_Comments='" & pcv_AdmComments & "' WHERE idOrder="&pcv_IdOrder&" AND pcPackageInfo_ID="&pcv_PackageID&";"
		end if
	else
		if pcv_shippedDate<>"" then
			if scDB="SQL" then
				query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "','" & dtShippedDate & "','" & pcv_tracking & "','" & pcv_AdmComments & "',1);"
			else
				query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "',#" & dtShippedDate & "#,'" & pcv_tracking & "','" & pcv_AdmComments & "',1);"
			end if
		else
				query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "','" & pcv_tracking & "','" & pcv_AdmComments & "',1);"
		end if
	end if
	set rs=connTemp.execute(query)
	set rs=nothing
	
	if pcv_PackageID=0 then
		query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_IdOrder & " ORDER by pcPackageInfo_ID DESC;"
		set rs=connTemp.execute(query)
		pcv_PackageID=rs("pcPackageInfo_ID")
		set rs=nothing
	end if
	qry_ID=pcv_IdOrder
	
	query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID & ";"
	set rstemp=connTemp.execute(query)
	query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier,pcPackageInfo_ID) VALUES (" & qry_ID & ",2,'" & pcv_AdmComments & "',0,0," & pcv_PackageID & ");"
	set rstemp=connTemp.execute(query)
	
	if intUSPSLabelFlag="1" then
		query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1 WHERE idOrder="&pcv_IdOrder&" AND pcPackageInfo_ID="&pcv_PackageID&";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	else
		if trim(pcv_PrdList)<>"" then
			pcA=split(pcv_PrdList,",")
			For i=lbound(pcA) to ubound(pcA)
				if trim(pcA(i)<>"") then
					query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1,pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND idProductOrdered=" & pcA(i) & ";"
					set rs=connTemp.execute(query)
					set rs=nothing
				end if
			Next
		else
			query="UPDATE ProductsOrdered SET pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND pcPrdOrd_Shipped=0 AND pcDropShipper_ID=0;"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
	end if
	
	pcv_SendCust="1"
	pcv_SendAdmin="0"
	
	pcv_LastShip="0"
	query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder AND ProductsOrdered.pcPrdOrd_Shipped=0) WHERE Orders.idorder=" & qry_ID & " AND Orders.orderstatus<>4;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_LastShip="0"
	else
		pcv_LastShip="1"
	end if
	set rs=nothing
	
	if intUSPSLabelFlag="1" then
		query="SELECT * FROM productsOrdered WHERE pcPrdOrd_Shipped<>1 AND idorder=" & qry_ID & ";"
		set rsQ=connTemp.execute(query)
		if rsQ.eof then
			query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
		else
			query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & qry_ID & ";"
		end if
		set rsQ=nothing
		set rs=connTemp.execute(query)
		set rs=nothing
	else
		if trim(pcv_PrdList)<>"" then
			if pcv_LastShip="1" then
				query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
			else
				query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & qry_ID & ";"
			end if
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	end if

	If pcv_LastShip="1" Then
		'// Perform a Google Action	
		pcv_strGoogleMethod = "mark" ' // Marks the order shipped at Google 			
		%> <!--#include file="../includes/GoogleCheckout_OrderManagement.asp"--> <%
	End If
	%>
	<!--#include file="../pc/inc_PartShipEmail.asp"-->
	<table class="pcCPcontent">
	<tr>
		<td valign="top">
		<table  border="0" cellpadding="0" cellspacing="0" width="60%">
		<tr>
			<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_IdOrder))%></b></td>
		</tr>
        <% if request("usps")="1" then %>
		<% else %>
		<tr>
			<td><b>Steps</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
			<td width="95%"><font color="#A8A8A8">Select products</font></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step2.gif"></td>
			<td><font color="#A8A8A8">Specify Shipment Details</font></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step3a.gif"></td>
			<td><b>Finalize Shipment</b></td>
		</tr>
        <% end if %>
		</table>
		</td>
	</tr>
	</table>
	<table class="pcCPcontent">
		<tr>
			<td align="center"><div class="pcCPmessageSuccess">Your order was updated successfully!</div></td>
		</tr>
		<tr>
			<td class="pcSpacer">&nbsp;</td>
		</tr>
		<tr>
			<td align="center">
				<a href="OrdDetails.asp?id=<%=pcv_IdOrder%>">Back to order details &gt;&gt;</a>
			</td>
		</tr>
	</table>
<%END IF%>
<%call closedb()%><!--#include file="AdminFooter.asp"-->