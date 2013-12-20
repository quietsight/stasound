<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true
Server.ScriptTimeout = 5400%>
<% pageTitle="Export Addresses to UPS" %>
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
<% dim rstemp, rstemp2, conntemp, mysql, strtext, File1

IF request("action")="add" then
	call opendb()
	File1=Server.MapPath("UPSAddresses.csv")
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
	
	strtext=""
	
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
		pcv_name=pcArray(1,i) & " " & pcArray(2,i)
		if InStr(pcv_name,",")>0 then
			pcv_name="""" & pcv_name & """"
		end if
		pcv_company=pcArray(3,i)
		if InStr(pcv_company,",")>0 then
			pcv_company="""" & pcv_company & """"
		end if
		pcv_phone=pcArray(4,i)
		if InStr(pcv_phone,",")>0 then
			pcv_phone="""" & pcv_phone & """"
		end if
		pcv_email=pcArray(5,i)
		pcv_addr=pcArray(6,i)
		pcv_addr2=pcArray(7,i)
		pcv_zip=pcArray(8,i)
		pcv_state=pcArray(9,i)
		if (pcv_state="") and (pcArray(19,i)<>"") then
			pcv_state=Left(pcArray(19,i),5)
		end if
		pcv_city=pcArray(10,i)
		if InStr(pcv_city,",")>0 then
			pcv_city="""" & pcv_city & """"
		end if
		pcv_country=pcArray(11,i)
		
		pcv_saddr=pcArray(12,i)
		pcv_scity=pcArray(13,i)
		if InStr(pcv_scity,",")>0 then
			pcv_scity="""" & pcv_scity & """"
		end if
		pcv_sstate=pcArray(14,i)
		if (pcv_sstate="") and (pcArray(20,i)<>"") then
			pcv_sstate=Left(pcArray(20,i),5)
		end if
		pcv_scountry=pcArray(15,i)
		pcv_szip=pcArray(16,i)
		pcv_scompany=pcArray(17,i)
		if InStr(pcv_scompany,",")>0 then
			pcv_scompany="""" & pcv_scompany & """"
		end if
		pcv_saddr2=pcArray(18,i)
		
		if pcv_company<>"" then
			pcv_tmp1=pcv_company
		else
			pcv_tmp1=pcv_name
		end if
		
		if len(pcv_addr)<=35 then
			if InStr(pcv_addr,",")>0 then
				pcv_tmpaddr1="""" & pcv_addr & """"
			else
				pcv_tmpaddr1=pcv_addr
			end if
			if InStr(pcv_addr2,",")>0 then
				pcv_tmpaddr2="""" & pcv_addr2 & """"
			else
				pcv_tmpaddr2=pcv_addr2
			end if
			pcv_tmp2=pcv_tmpaddr1 & "," & pcv_tmpaddr2 & ",,"
		else
			pcv_addr=trim(pcv_addr)
			k=len(pcv_addr)
			Do
				k=InStrRev(k-1,pcv_addr," ")
			Loop Until k<35
			A=trim(mid(pcv_addr,1,k-1))
			B=trim(mid(pcv_addr,k+1,len(pcv_addr)))
			if InStr(A,",")>0 then
				A="""" & A & """"
			end if
			if InStr(B,",")>0 then
				B="""" & B & """"
			end if
			if InStr(pcv_addr2,",")>0 then
				pcv_tmpaddr2="""" & pcv_addr2 & """"
			else
				pcv_tmpaddr2=pcv_addr2
			end if
			pcv_tmp2=A & "," & B & "," & pcv_tmpaddr2 & ","
		end if
		if trim(pcv_addr)<>"" then
			strtext=strtext & pcv_tmp1 & "," & pcv_name & "," & pcv_name & "," & pcv_tmp2 & pcv_city & "," & pcv_state & "," & pcv_zip & "," & pcv_country & "," & pcv_phone & ",," & pcv_email & ",,0,,0" & vbcrlf
		end if
		
		if (pcv_saddr<>"") and (pcv_saddr<>pcv_addr) then
		
			if pcv_scompany<>"" then
				pcv_tmp1=pcv_scompany
			else
				pcv_tmp1=pcv_name
			end if
		
			if len(pcv_saddr)<=35 then
				if InStr(pcv_saddr,",")>0 then
				pcv_tmpaddr1="""" & pcv_saddr & """"
				else
				pcv_tmpaddr1=pcv_saddr
				end if
				if InStr(pcv_saddr2,",")>0 then
				pcv_tmpaddr2="""" & pcv_saddr2 & """"
				else
				pcv_tmpaddr2=pcv_saddr2
				end if
				pcv_tmp2=pcv_tmpaddr1 & "," & pcv_tmpaddr2 & ",,"
			else
				pcv_saddr=trim(pcv_saddr)
				k=len(pcv_saddr)
				Do
					k=InStrRev(k-1,pcv_saddr," ")
				Loop Until k<35
				A=trim(mid(pcv_saddr,1,k-1))
				B=trim(mid(pcv_saddr,k+1,len(pcv_saddr)))
				if InStr(A,",")>0 then
				A="""" & A & """"
				end if
				if InStr(B,",")>0 then
				B="""" & B & """"
				end if
				if InStr(pcv_saddr2,",")>0 then
				pcv_tmpaddr2="""" & pcv_saddr2 & """"
				else
				pcv_tmpaddr2=pcv_saddr2
				end if
				pcv_tmp2=A & "," & B & "," & pcv_tmpaddr2 & ","
			end if
			if trim(pcv_saddr)<>"" then
			strtext=strtext & pcv_tmp1 & "," & pcv_name & "," & pcv_name & "," & pcv_tmp2 & pcv_scity & "," & pcv_sstate & "," & pcv_szip & "," & pcv_scountry & "," & pcv_phone & ",," & pcv_email & ",,0,,0" & vbcrlf
			end if
		end if
		
		query="SELECT recipient_FullName, recipient_Address, recipient_City, recipient_StateCode, recipient_Zip, recipient_CountryCode, recipient_Company, recipient_Address2, recipient_State FROM recipients where idcustomer=" & pcv_idcust & ";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		IF not rs.eof THEN
			pcArray1=rs.getRows()
			RCount1 = UBound(pcArray1, 2)
			set rs=nothing
		
			For j=0 to RCount1
				pcv_fname=pcArray1(0,j)
				if InStr(pcv_fname,",")>0 then
					pcv_fname="""" & pcv_fname & """"
				end if
				pcv_raddr=pcArray1(1,j)
				pcv_rcity=pcArray1(2,j)
				if InStr(pcv_rcity,",")>0 then
					pcv_rcity="""" & pcv_rcity & """"
				end if
				pcv_rstate=pcArray1(3,j)
				if (pcv_rstate="") and (pcArray1(8,j)<>"") then
					pcv_rstate=left(pcArray1(8,j),5)
				end if
				pcv_rzip=pcArray1(4,j)
				if Instr(pcv_rzip,"||")>0 then
					A=split(pcv_rzip,"||")
					pcv_rzip=A(0)
					pcv_rphone=A(1)
					if InStr(pcv_rphone,",")>0 then
					pcv_rphone="""" & pcv_rphone & """"
					end if
				end if
				pcv_rcountry=pcArray1(5,j)
				pcv_rcompany=pcArray1(6,j)
				if InStr(pcv_rcompany,",")>0 then
					pcv_rcompany="""" & pcv_rcompany & """"
				end if
				pcv_raddr2=pcArray1(7,j)
				
				if (pcv_raddr<>pcv_addr) and (trim(pcv_fname)<>"") then
				
					if pcv_rcompany<>"" then
						pcv_tmp1=pcv_rcompany
					else
						pcv_tmp1=pcv_fname
					end if
			
					if len(pcv_raddr)<=35 then
						if InStr(pcv_raddr,",")>0 then
							pcv_tmpaddr1="""" & pcv_raddr & """"
						else
							pcv_tmpaddr1=pcv_raddr
						end if
						if InStr(pcv_raddr2,",")>0 then
							pcv_tmpaddr2="""" & pcv_raddr2 & """"
						else
							pcv_tmpaddr2=pcv_raddr2
						end if
						pcv_tmp2=pcv_tmpaddr1 & "," & pcv_tmpaddr2 & ",,"
					else
						pcv_raddr=trim(pcv_raddr)
						k=len(pcv_raddr)
						Do
							k=InStrRev(k-1,pcv_raddr," ")
						Loop Until k<35
						A=trim(mid(pcv_raddr,1,k-1))
						B=trim(mid(pcv_raddr,k+1,len(pcv_raddr)))
						if InStr(A,",")>0 then
							A="""" & A & """"
						end if
						if InStr(B,",")>0 then
							B="""" & B & """"
						end if
						if InStr(pcv_raddr2,",")>0 then
							pcv_tmpaddr2="""" & pcv_raddr2 & """"
						else
							pcv_tmpaddr2=pcv_raddr2
						end if
						pcv_tmp2=A & "," & B & "," & pcv_tmpaddr2 & ","
					end if
					if trim(pcv_raddr)<>"" then
						strtext=strtext & pcv_tmp1 & "," & pcv_fname & "," & pcv_fname & "," & pcv_tmp2 & pcv_rcity & "," & pcv_rstate & "," & pcv_rzip & "," & pcv_rcountry & "," & pcv_rphone & ",,,,0,,0" & vbcrlf
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
			<p>Customer addresses were exported successfully	to the file: <a href="UPSAddresses.csv">UPSAddresses.csv</a>.</p>
<p>To download the file, either right-click on the file name and select '<strong>Save Target As...</strong>' or FTP into the <%=scAdminFolderName%> folder and download it.</p>
<p>When you import the file into your Address Books at UPS.com, select &quot;<u>My UPS Address Book</u>&quot; from the &quot;<u>Original File Format</u>&quot; drop-down menu.
				<br>
				<br>
				<br>
				
<input type="button" value="Back" onclick="location='exportData.asp#ups'" class="ibtnGrey">
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