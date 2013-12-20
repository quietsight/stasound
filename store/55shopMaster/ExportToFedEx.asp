<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true
Server.ScriptTimeout = 5400%>
<% pageTitle="Export to FedEx" %>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/FedEXconstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<% dim rstemp, rstemp2, conntemp, mysql, strtext, File1

IF request("action")="add" then
	call opendb()

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
	strTDateVar3=Request.Form("ShipDate")
	DateVar3=strTDateVar3
	if (strTDateVar3<>"") and (isDate(strTDateVar3)) then
		if scDateFrmt="DD/MM/YY" then
			DateVarArray3=split(strTDateVar3,"/")
			DateVar3=(DateVarArray3(2) & "-" & DateVarArray3(1)&"-"&DateVarArray3(0))
		else
			DateVarArray3=split(strTDateVar3,"/")
			DateVar3=(DateVarArray3(2) & "-" & DateVarArray3(0)&"-"&DateVarArray3(1))
		end if
	end if
	if DateVar3="" then
		DateVar3=Year(Date()) & "-" & Month(Date()) & "-" & Day(Date())
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
	
	query3=""
	If (request("otype")<>"") AND (request("otype")<>"0") then
		query3=" AND OrderStatus=" & request("otype")
	End If
	
	query="SELECT DISTINCT orders.idCustomer, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_ShippingPhone, orders.ShippingFullName, orders.shippingCompany, orders.shippingAddress2, orders.pcOrd_ShippingEmail, orders.pcOrd_ShipWeight, orders.shipmentDetails, orders.ordShiptype, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.address2, customers.customerCompany, customers.phone, customers.email, customers.name, customers.lastName FROM orders INNER JOIN customers ON orders.idCustomer = customers.idcustomer  WHERE idOrder>0 " & query1 & query2 & query3 & ";"
	
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if rs.eof then
		set rs=nothing
		call closeDb()
		response.Redirect "msg.asp?message=46"
	end if
	
	call opendb()
	tmpFN="FedEx" & Month(Date()) & Day(Date()) & Year(Date()) & Hour(Time()) & Minute(Time()) & Second(Time()) & ".csv"
	File1=Server.MapPath(tmpFN)
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set F=fso.CreateTextFile(File1,True)
	
	F.WriteLine("""RECIPIENT ID"",""COUNTRY"",""CONTACT NAME"",""COMPANY NAME"",""ADDRESS 1"",""ADDRESS 2"",""POSTAL CODE"",""STATE"",""CITY"",""PHONE"",""RESIDENTIAL FLAG"",""WEIGHT"",""SERVICE TYPE"",""PACKAGE TYPE"",""SHIP DATE"",""BILL TO"",""RECIPIENT EMAIL ADDRESS"",")

	pcArray=rs.getRows()
	RCount = UBound(pcArray, 2)

	 
	set rs=nothing
	
	For i=0 to RCount
		sbo_IdCustomer = pcArray(0,i)
		sbo_shippingAddress = pcArray(1,i)
		sbo_shippingStateCode = pcArray(2,i)
		sbo_shippingState = pcArray(3,i)
		sbo_shippingCity = pcArray(4,i)
		sbo_shippingCountryCode = pcArray(5,i)
		sbo_shippingZip = pcArray(6,i)
		sbo_pcOrd_ShippingPhone = pcArray(7,i)
		sbo_ShippingFullName = pcArray(8,i)
		sbo_shippingCompany = pcArray(9,i)
		sbo_shippingAddress2 = pcArray(10,i)
		sbo_pcOrd_ShippingEmail = pcArray(11,i)
		sbo_pcOrd_ShipWeight = pcArray(12,i)
		sbo_shipmentDetails = pcArray(13,i)
		sbo_ordShiptype = pcArray(14,i)
		sbo_address = pcArray(15,i)
		sbo_zip = pcArray(16,i)
		sbo_stateCode = pcArray(17,i)
		sbo_state = pcArray(18,i)
		sbo_city = pcArray(19,i)
		sbo_countryCode = pcArray(20,i)
		sbo_address2 = pcArray(21,i)
		sbo_customerCompany = pcArray(22,i)
		sbo_phone = pcArray(23,i)
		sbo_email = pcArray(24,i)
		sbo_name = pcArray(25,i)
		sbo_lastName = pcArray(26,i)
		If TRIM(sbo_shippingAddress)&""="" Then
			sbo_shippingAddress = sbo_address
			sbo_shippingStateCode = sbo_stateCode
			sbo_shippingState = sbo_state
			sbo_shippingCity = sbo_city
			sbo_shippingCountryCode = sbo_countryCode
			sbo_shippingZip = sbo_zip
			sbo_shippingAddress2 = sbo_address2
		End If
		If TRIM(sbo_pcOrd_ShippingPhone)&""="" Then
			sbo_pcOrd_ShippingPhone = sbo_phone
		End If
		If TRIM(sbo_ShippingFullName)&""="" Then
			sbo_ShippingFullName = sbo_name&" "&sbo_lastName
		End If
		If TRIM(sbo_shippingCompany)&""="" Then
			sbo_shippingCompany = sbo_customerCompany
		End If
		If TRIM(sbo_pcOrd_ShippingEmail)&""="" Then
			sbo_pcOrd_ShippingEmail = sbo_email
		End If
		
		set StringBuilderObj = new StringBuilder
		
		StringBuilderObj.append chr(34) & sbo_IdCustomer & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_shippingCountryCode & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_ShippingFullName & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_shippingCompany & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_shippingAddress & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_shippingAddress2 & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_shippingZip & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_shippingStateCode & sbo_shippingState & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_shippingCity & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_pcOrd_ShippingPhone & chr(34) & ","
		'// RESIDENTIAL FLAG
		if sbo_ordShiptype="1" then
			tmpRes="0"
		else
			tmpRes="1"
		end if
		StringBuilderObj.append chr(34) & tmpRes & chr(34) & ","
		
		'//Order Weight
		tWeight=sbo_pcOrd_ShipWeight
		
		if scShipFromWeightUnit="KGS" then
			pKilos=Int(tWeight/1000)
			pWeight_g=tWeight-(pKilos*1000)
			oWeight_kg=pKilos
			if pWeight_g>0 then
				oWeight_kg=oWeight_kg+1
			end if
			StringBuilderObj.append chr(34) & oWeight_kg & chr(34) & ","
		else
			pPounds=Int(tWeight/16)
			pWeight_oz=tWeight-(pPounds*16)
			oWeight=pPounds
			if pWeight_oz>0 then
				oWeight=oWeight+1
			end if
			StringBuilderObj.append chr(34) & oWeight & chr(34) & ","
		end if
		
		'//Service Type
		tmpSerType=0
		pshipmentDetails=sbo_shipmentDetails
		if Instr(pshipmentDetails,",")>0 then
			shipping=split(pshipmentDetails,",")
			if ubound(shipping)>1 then
				if NOT isNumeric(trim(shipping(2))) then
				else
					Shipper=shipping(0)
					Service=shipping(5)
					Select Case trim(Service)
						Case "FIRST_OVERNIGHT": tmpSerType="06"
						Case "PRIORITY_OVERNIGHT": tmpSerType="01"
						Case "STANDARD_OVERNIGHT": tmpSerType="05"
						Case "FEDEX_2_DAY": tmpSerType="03"
						Case "FEDEX_EXPRESS_SAVER": tmpSerType="20"
						Case "FEDEX_GROUND": tmpSerType="92"
						Case "GROUND_HOME_DELIVERY": tmpSerType="90"
					End Select
				end if
			end if
		end if
		StringBuilderObj.append chr(34) & tmpSerType & chr(34) & ","
		
		'//Package Type
		tmpPackType="01"
		Select Case FEDEX_FEDEX_PACKAGE
			Case "YOURPACKAGING": tmpPackType="01"
			Case "FEDEX10KGBOX": tmpPackType="03"
			Case "FEDEX25KGBOX": tmpPackType="03"
			Case "FEDEXBOX": tmpPackType="03"
			Case "FEDEXENVELOPE": tmpPackType="06"
			Case "FEDEXPAK": tmpPackType="02"
			Case "FEDEXTUBE": tmpPackType="04"
		End Select
		StringBuilderObj.append chr(34) & tmpPackType & chr(34) & ","
		
		StringBuilderObj.append chr(34) & DateVar3 & chr(34) & ","
		StringBuilderObj.append chr(34) & "1" & chr(34) & ","
		StringBuilderObj.append chr(34) & sbo_pcOrd_ShippingEmail & chr(34) & ","
		
		F.WriteLine(StringBuilderObj.toString())
		set StringBuilderObj = nothing
		
		sbo_IdCustomer = ""
		sbo_shippingAddress = ""
		sbo_shippingStateCode = ""
		sbo_shippingState = ""
		sbo_shippingCity = ""
		sbo_shippingCountryCode = ""
		sbo_shippingZip = ""
		sbo_pcOrd_ShippingPhone = ""
		sbo_ShippingFullName = ""
		sbo_shippingCompany = ""
		sbo_shippingAddress2 = ""
		sbo_pcOrd_ShippingEmail = ""
		sbo_pcOrd_ShipWeight = ""
		sbo_shipmentDetails = ""
		sbo_ordShiptype = ""
		sbo_address = ""
		sbo_zip = ""
		sbo_stateCode = ""
		sbo_state = ""
		sbo_city = ""
		sbo_countryCode = ""
		sbo_address2 = ""
		sbo_customerCompany = ""
		sbo_phone = ""
		sbo_email = ""
		sbo_name = ""
		sbo_lastName = ""
	Next
		
	F.Write(strtext)
	
	F.Close
	Set fso=Nothing
	%>
<table width="94%" border="0" cellspacing="0" cellpadding="4" align="center">
<tr> 
			
<td class="normal">
	<div class="pcCPmessageSuccess">
		<p>Selected Orders were exported successfully to the file: <a href="<%=tmpFN%>"><%=tmpFN%></a>.</p>
		<p>To download the file, either right-click on the file name and select '<strong>Save Target As...</strong>' or FTP into the <%=scAdminFolderName%> folder and download it.</p>
	</div>
				<br>
				<br>
				<br>
				
<input type="button" value="Back" onclick="location='exportData.asp#fedex'" class="ibtnGrey">
				&nbsp;
<input type="button" value="Start Page" onclick="location='menu.asp'" class="ibtnGrey">
				<br>
				<br>
				</p>
<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
			</td>
		</tr>
	</table>
<% END IF %>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->