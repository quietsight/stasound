<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp" --> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="../includes/GoogleCheckout_APIFunctions.asp"-->
<!--#include file="pcPay_GoogleCheckout_Handler.asp"-->
<!--#include file="Header.asp"-->
<%
IF request("action")="add" THEN
	Dim connTemp,rs,query
	call opendb()
	pcv_IdOrder=getUserInput(request("idorder"),0)
	if ((pcv_IdOrder="") OR (not validNum(pcv_IdOrder))) then
		pcv_IdOrder=0
	end if
	pcv_PrdList=getUserInput(request("PrdList"),0)

	pcv_count=request("count")
	if pcv_count="" then
		pcv_count=0
	end if

	if (pcv_IdOrder=0) then
		response.redirect "menu.asp"
	end if
	
	pcv_method=getUserInput(request("pcv_method"),0)
	pcv_tracking=getUserInput(request("pcv_tracking"),0)
	pcv_shippedDate=getUserInput(request("pcv_shippedDate"),0)
	pcv_AdmComments=getUserInput(request("pcv_AdmComments"),0)
	if pcv_AdmComments<>"" then
		pcv_AdmComments=replace(pcv_AdmComments,"'","''")
	end if
	
	query="SELECT orders.pcOrd_GoogleIDOrder FROM orders WHERE idOrder="& pcv_IdOrder
	set rs=server.CreateObject("ADODB.RecordSet")
	Set rs=conntemp.execute(query)
	if Not rs.eof then
		pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder")
		pIdOrder=pcv_IdOrder
		pcv_tempTrackingNumber=pcv_tracking
	end if
	set rs=nothing

	
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
	if pcv_shippedDate<>"" then
		if scDB="SQL" then
			query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "','" & dtShippedDate & "','" & pcv_tracking & "','" & pcv_AdmComments & "',1);"
		else
			query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "',#" & dtShippedDate & "#,'" & pcv_tracking & "','" & pcv_AdmComments & "', 1);"
		end if
	else
		query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) VALUES (" & pcv_IdOrder & ",'" & pcv_method & "','" & pcv_tracking & "','" & pcv_AdmComments & "', 1);"
	end if
	set rs=connTemp.execute(query)
	set rs=nothing
	
	query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_IdOrder & " ORDER by pcPackageInfo_ID DESC;"
	set rs=connTemp.execute(query)
	pcv_PackageID=rs("pcPackageInfo_ID")
	set rs=nothing
	
	qry_ID=pcv_IdOrder
	
	query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID & ";"
	set rstemp=connTemp.execute(query)
	query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier,pcPackageInfo_ID) VALUES (" & qry_ID & ",2,'" & pcv_AdmComments & "'," & session("pc_idsds") & "," & session("pc_sdsIsDropShipper") & "," & pcv_PackageID & ");"
	set rstemp=connTemp.execute(query)

	'+++++++++++++++++++++++++++++++++++
	if trim(pcv_PrdList)<>"" then
		pcA=split(pcv_PrdList,",")
		For i=lbound(pcA) to ubound(pcA)
			if trim(pcA(i)<>"") then
				query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1, pcPackageInfo_ID=" & pcv_PackageID & " WHERE (idorder=" & qry_ID & " AND idProduct=" & pcA(i) & ");"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
		Next
	else
		query="UPDATE ProductsOrdered SET pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND pcPrdOrd_Shipped=0 AND pcDropShipper_ID=0;"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	end if
	'+++++++++++++++++++++++++++++++++++
	
	
	if session("pc_sdsIsDropShipper")="1" then
		query="SELECT pcSupplier_CustNotifyUpdates As A,pcSupplier_Company As B,pcSupplier_FirstName As C,pcSupplier_LastName As D,pcSupplier_Email As E FROM pcSuppliers WHERE pcSupplier_ID=" & session("pc_idsds") & ";"
	else
		query="SELECT pcDropShipper_CustNotifyUpdates As A,pcDropShipper_Company As B,pcDropShipper_FirstName As C,pcDropShipper_LastName As D,pcDropShipper_Email As E FROM pcDropShippers WHERE pcDropShipper_ID=" & session("pc_idsds") & ";"
	end if
	Set rs=connTemp.execute(query)
	pcv_CustNotice=0
	pcv_UseDropShipperInfo=0
	if not rs.eof then
		pcv_CustNotice=rs("A")
		if IsNull(pcv_CustNotice) or pcv_CustNotice="" then
			pcv_CustNotice=0
		end if
		pcv_UseDropShipperInfo=1
		pcv_DS_Name=rs("B") & " (" & rs("C") & " " & rs("D") & ")"
		pcv_DS_Email=rs("E")
	end if
	set rs=nothing
	
	if pcv_CustNotice="1" then
		pcv_SendCust="1"
	else
		pcv_SendCust="0"
	end if
	
	pcv_SendAdmin="1"
	pcv_LastShipDS="0"
	pcv_LastShip="0"

	'// Update order status and send e-mails
	
	'// Update drop-shipper specific order status
		query="SELECT idproduct FROM ProductsOrdered WHERE pcPrdOrd_Shipped=1 AND pcDropShipper_ID=" & session("pc_idsds") & " AND idorder=" & qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		productCount = 0
		if NOT rs.eof then					
			do while NOT rs.eof
				productCount = productCount + 1
			rs.movenext
			loop
		end if
		set rs=nothing
				
		if int(productCount) < int(pcv_count) then
			pcv_LastShipDS="0"
		else
			pcv_LastShipDS="1"
		end if
			
		'// Check to see if a record for this order already exists
		query="SELECT pcDropShipO_ID FROM pcDropShippersOrders WHERE pcDropShipO_idOrder=" & qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if rs.eof then
			pcv_orderExists="0"
		else
			pcv_orderExists="1"
		end if
		set rs=nothing

		if trim(pcv_PrdList)<>"" then
			if pcv_orderExists="1" then
				if pcv_LastShipDS="1" then
					query="UPDATE pcDropShippersOrders SET pcDropShipO_OrderStatus=4 WHERE pcDropShipO_idOrder=" & qry_ID & ";"
				else
					query="UPDATE pcDropShippersOrders SET pcDropShipO_OrderStatus=7 WHERE pcDropShipO_idOrder=" & qry_ID & ";"
				end if			
			else
				if pcv_LastShipDS="1" then
					query="INSERT INTO pcDropShippersOrders (pcDropShipO_DropShipper_ID,pcDropShipO_idOrder,pcDropShipO_OrderStatus) VALUES (" & session("pc_idsds") & "," & qry_ID & ",4);"
				else
					query="INSERT INTO pcDropShippersOrders (pcDropShipO_DropShipper_ID,pcDropShipO_idOrder,pcDropShipO_OrderStatus) VALUES (" & session("pc_idsds") & "," & qry_ID & ",7);"
				end if
			end if
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	
	'// Update general order status and send e-mails
	query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder AND ProductsOrdered.pcPrdOrd_Shipped=0) WHERE Orders.idorder=" & qry_ID & " AND Orders.orderstatus<>4;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcv_LastShip="0"
	else
		pcv_LastShip="1"
	end if
	set rs=nothing

	if trim(pcv_PrdList)<>"" then
		if pcv_LastShip="1" then
			query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
		else
			query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & qry_ID & ";"
		end if
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	
	If pcv_LastShip="1" Then
		'// Perform a Google Action	
		pcv_strGoogleMethod = "mark" ' // Marks the order shipped at Google 	
		%> <!--#include file="../includes/GoogleCheckout_OrderManagement.asp"--> <%
	End If
	%>
	<!--#include file="inc_PartShipEmail.asp"-->
<div id="pcMain">
	<table class="pcMainTable">
	<tr>
		<td valign="top">
		<table class="pcShowContent">
		<tr>
			<td colspan="6"><h1><%response.write dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%> - <%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_1")%> <%=(scpre+int(pcv_IdOrder))%></h1></td>
		</tr>
		<tr>
			<td colspan="6" class="pcSpacer"></td>
		</tr>
		<tr>
			<td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
			<td width="28%" nowrap><font color="#A8A8A8"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_3")%></font></td>
			<td width="5%" align="center"><img border="0" src="images/step2.gif"></td>
			<td width="28%" nowrap><font color="#A8A8A8"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_4")%></font></td>
			<td width="5%" align="center"><img border="0" src="images/step3a.gif"></td>
			<td width="29%" nowrap><b><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_5")%></b></td>
		</tr>
		<tr>
			<td colspan="6" class="pcSpacer"></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
	<table class="pcMainTable">
		<tr>
			<td><div class="pcSuccessMessage"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_20")%></div></td>
		</tr>
		<tr>
			<td class="pcSpacer">&nbsp;</td>
		</tr>
		<tr>
			<td>
				<a href="sds_viewPastD.asp?idOrder=<%response.write (scpre+int(pcv_IdOrder))%>"><img src="<%=rslayout("pcLO_backtoOrder")%>" border="0"></a>
			</td>
		</tr>
	</table>
</div>
<%call closedb()%>
<%
ELSE
	response.Redirect "sds_ViewPast.asp"
	response.End()
END IF
%>
<!--#include file="Footer.asp"-->