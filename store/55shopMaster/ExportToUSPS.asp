<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true
Server.ScriptTimeout = 5400%>
<% pageTitle="Export Addresses to USPS" %>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="frooglecurrencyformatinc.asp"--> 
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/USPSCountry.asp"-->

<%
dim rs, conntemp, query, fso, f, pcv_ExportType, pcv_FromDate, pcv_ToDate, File1

IF request("action")="add" then
	call opendb()
	File1=Server.MapPath("USPSAddresses.csv")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set F=fso.CreateTextFile(File1,True)

	pcv_ExportType=request("exporttype")
	pcv_FromDate=request("FromDate")
	pcv_ToDate=request("ToDate")
	
	if not IsDate(pcv_FromDate) then
		pcv_FromDate=""
	end if
	
	if not IsDate(pcv_ToDate) then
		pcv_ToDate=""
	end if
	
	strtext=strtext & "Full Name,Company,Address 1,Address 2,Address 3,City,State,Zip Code,Province,Country,Urbanization,Phone,Fax,E Mail,Customer Ref #,Short Name" & vbcrlf
	
	IF pcv_ExportType="0" then
		query="SELECT idcustomer, name, lastName, customerCompany, phone, email, address, address2, zip, stateCode, city, countryCode, shippingaddress, shippingcity, shippingStateCode, shippingCountryCode, shippingZip, shippingCompany, shippingAddress2, state, shippingState FROM customers where customerType<2 AND suspend=0"
	Else
		query1=""
		if pcv_FromDate<>"" then
			if scDB="Access" then
				query1=query1 & " AND orders.orderDate>=#" & pcv_FromDate & "# "
			else
				query1=query1 & " AND orders.orderDate>='" & pcv_FromDate & "' "
			end if
		end if
		if pcv_ToDate<>"" then
			if scDB="Access" then
				query1=query1 & " AND orders.orderDate<=#" & pcv_ToDate & "# "
			else
				query1=query1 & " AND orders.orderDate<='" & pcv_ToDate & "' "
			end if
		end if
		query="SELECT orders.idCustomer, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email, customers.address, customers.address2, customers.zip, customers.stateCode, customers.city, customers.countryCode, customers.shippingaddress, customers.shippingcity, customers.shippingStateCode, customers.shippingCountryCode, customers.shippingZip, customers.shippingCompany, customers.shippingAddress2, customers.state, customers.shippingState FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.orderStatus)>1 And (orders.orderStatus)<5) " & query1 & ") OR (((orders.orderStatus)>6 And (orders.orderStatus)<9) " & query1 & ") OR (((orders.orderStatus)=10 Or (orders.orderStatus)=12) " & query1 & ") GROUP BY orders.idCustomer, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email, customers.address, customers.address2, customers.zip, customers.stateCode, customers.city, customers.countryCode, customers.shippingaddress, customers.shippingcity, customers.shippingStateCode, customers.shippingCountryCode, customers.shippingZip, customers.shippingCompany, customers.shippingAddress2, customers.state, customers.shippingState;"
		
	End If

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if rs.eof then
		set rs=nothing
		call closeDb()
		response.Redirect "msg.asp?message=43"
	end if
	
	pcArray=rs.getRows()
	RCount = UBound(pcArray, 2)
	set rs=nothing
	
	For i=0 to RCount
		pcv_idcust=pcArray(0,i)
		
		'// Billing Address
		pcv_name=pcArray(1,i) & " " & pcArray(2,i)
		
		pcv_company=pcArray(3,i)
		pcv_company=replace(pcv_company,","," ")
		
		pcv_phone=pcArray(4,i)
		pcv_phone=replace(pcv_phone,",","")
		pcv_phone=replace(pcv_phone,"(","")
		pcv_phone=replace(pcv_phone,")","")
		pcv_phone=replace(pcv_phone,"-","")
		pcv_phone=replace(pcv_phone,".","")
		pcv_phone=replace(pcv_phone," ","")
			
		pcv_email=pcArray(5,i)
		pcv_addr=pcArray(6,i)
		pcv_addr=replace(pcv_addr,",","")
		pcv_addr2=pcArray(7,i)
		pcv_addr2=replace(pcv_addr2,",","")

		pcv_zip=pcArray(8,i)
		
		pcv_state=pcArray(9,i)
		if (pcv_state="") and (pcArray(19,i)<>"") then
			pcv_state=Left(pcArray(19,i),5)
		end if
		
		pcv_city=pcArray(10,i)
		pcv_city=replace(pcv_city,",","")

		pcv_country=pcArray(11,i)
		if ucase(pcv_country)="US" then
			pcv_country="United States"
		else
			pcv_country=USPSCountry(pcv_country)
		end if

		'// Shipping Address
		pcv_saddr=pcArray(12,i)
		if not isNull(pcv_saddr) and pcv_saddr<>"" then
			pcv_saddr=replace(pcv_saddr,",","")
		end if
		
		pcv_saddr2=pcArray(18,i)
		if not isNull(pcv_saddr2) and pcv_saddr2<>"" then
			pcv_saddr2=replace(pcv_saddr2,",","")
		end if
		
		pcv_scity=pcArray(13,i)
		if not isNull(pcv_scity) and pcv_scity<>"" then
			pcv_scity=replace(pcv_scity,",","")
		end if
		
		pcv_sstate=pcArray(14,i)
		if (pcv_sstate="") and (pcArray(20,i)<>"") then
			pcv_sstate=Left(pcArray(20,i),5)
		end if
		
		pcv_scountry=pcArray(15,i)
		if ucase(pcv_scountry)="US" then
			pcv_scountry="United States"
		else
			pcv_scountry=USPSCountry(pcv_scountry)
		end if
		
		pcv_szip=pcArray(16,i)
		
		pcv_scompany=pcArray(17,i)
		if not isNull(pcv_scompany) and pcv_scompany<>"" then
			pcv_scompany=replace(pcv_scompany,",","")
		end if
		
		'Import Billing Address
		if trim(pcv_addr)<>"" then
			strtext=strtext & pcv_name & "," & pcv_company & "," & pcv_addr & "," & pcv_addr2 & "," & Address3 & "," & pcv_city & "," & pcv_state & "," & pcv_zip & "," & Province & "," & pcv_country & "," & Urbanization & "," & pcv_phone & "," & Fax & "," & pcv_email & "," & pcv_idcust & "," & ShortName & vbcrlf

		end if
		
		'Import Shipping Address
		if (pcv_saddr<>"") and (pcv_saddr<>pcv_addr) then
			strtext=strtext & pcv_name & "," & pcv_company & "," & pcv_saddr & "," & pcv_saddr2 & "," & Address3 & "," & pcv_scity & "," & pcv_sstate & "," & pcv_szip & "," & Province & "," & pcv_scountry & "," & Urbanization & "," & pcv_phone & "," & Fax & "," & pcv_email & "," & pcv_idcust & "," & ShortName & vbcrlf
		end if
		
		'Import all recipient addresses
		query="SELECT recipient_FullName, recipient_Address, recipient_City, recipient_StateCode, recipient_Zip, recipient_CountryCode, recipient_Company, recipient_Address2, recipient_State FROM recipients where idcustomer=" & pcv_idcust & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
			
		IF not rs.eof THEN
			pcArray1=rs.getRows()
			RCount1 = UBound(pcArray1, 2)
			set rs=nothing
			
			For j=0 to RCount1
				pcv_fname=pcArray1(0,j)
				pcv_fname=replace(pcv_fname,",","")
				
				pcv_raddr=pcArray1(1,j)
				pcv_rcity=pcArray1(2,j)
				pcv_rcity=replace(pcv_rcity,",","")

				pcv_rstate=pcArray1(3,j)
				if (pcv_rstate="") and (pcArray1(8,j)<>"") then
					pcv_rstate=left(pcArray1(8,j),5)
				end if
					
				pcv_rzip=pcArray1(4,j)
				if Instr(pcv_rzip,"||")>0 then
					A=split(pcv_rzip,"||")
					pcv_rzip=A(0)
					pcv_rphone=A(1)
					pcv_rphone=replace(pcv_rphone,",","")
					pcv_rphone=replace(pcv_rphone,"(","")
					pcv_rphone=replace(pcv_rphone,")","")
					pcv_rphone=replace(pcv_rphone,"-","")
					pcv_rphone=replace(pcv_rphone,".","")
					pcv_rphone=replace(pcv_rphone," ","")
				end if
				pcv_rcountry=pcArray1(5,j)
				if ucase(pcv_rcountry)="US" then
					pcv_rcountry="United States"
				else
					pcv_rcountry=USPSCountry(pcv_rcountry)
				end if
				pcv_rcompany=pcArray1(6,j)
				pcv_rcompany=replace(pcv_rcompany,",","")
				pcv_raddr2=pcArray1(7,j)
				
				if (pcv_raddr<>pcv_addr) and (trim(pcv_fname)<>"") then
			
					pcv_raddr=replace(pcv_raddr,",","")
					pcv_raddr2=replace(pcv_raddr2,",","")
					
					if trim(pcv_raddr)<>"" then
						strtext=strtext & pcv_fname & "," & pcv_rcompany & "," & pcv_raddr & "," & pcv_raddr2 & "," & Address3 & "," & pcv_rcity & "," & pcv_rstate & "," & pcv_rzip & "," & Province & "," & pcv_rcountry & "," & Urbanization & "," & pcv_rphone & "," & Fax & "," & pcv_email & "," & pcv_idcust & "," & ShortName & vbcrlf
					end if
				end if
			Next
		END IF
	
	Next

	F.Write(strtext)
	
	F.Close
	Set fso=Nothing
	%>
	<table width="94%" border="0" cellspacing="0" cellpadding="4" align="center">
		<tr> 
			
<td class="normal">
			<p>Customer addresses were exported successfully	to the file: <a href="USPSAddresses.csv">USPSAddresses.csv</a>.</p>
<p>To download the file, either right-click on the file name and select '<strong>Save Target As...</strong>' or FTP into the <%=scAdminFolderName%> folder and download it.</p>
<p>You can then import the file directly into your USPS Shipping Assistant. The USPS Shipping Assistant will validate all addresses at time of import and will generate a log file of any addresses that were not imported due to invalid data.
				<br>
				<br>
				<br>
				
<input type="button" value="Back" onclick="location='exportData.asp#USPS'" class="ibtnGrey">
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
<%call closedb()%><!--#include file="AdminFooter.asp"-->