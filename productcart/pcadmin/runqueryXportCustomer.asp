<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Customer Information" %>
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
<!--#include file="../includes/SQLFormat.txt"-->
<% 
response.Buffer=true
Response.Expires=0
Dim FieldList(100,2)
dim mySQL, conntemp, rstemp

call openDb()
' Choose the records to display
idcustomer=request.form("idcustomer")
pname=request.form("name")
lastName=request.form("lastName")
customerCompany=request.form("customerCompany")
phone=request.form("phone")
email=request.form("email")
address=request.form("address")
address2=request.form("address2")
city=request.form("city")
stateCode=request.form("stateCode")
zip=request.form("zip")
CountryCode=request.form("CountryCode")
customerType=request.form("customerType")
pcvrp_accrued=request.form("pcrp_accrued")
pcvrp_used=request.form("pcrp_used")
pcvrp_available=request.Form("pcrp_available")
fieldCount=request.Form("FieldCount")
For i=1 to fieldCount
	FieldList(i,0)=trim(request("pccust_cf" & i))
	if FieldList(i,0)<>"" then
		query="SELECT pcCField_Name FROM pcCustomerFields WHERE pcCField_ID=" & FieldList(i,0)
		set rs=connTemp.execute(query)
		if not rs.eof then
			FieldList(i,1)=rs("pcCField_Name")
		end if
		set rs=nothing
	end if
Next
pcvcust_recvnews=request.form("pccust_recvnews")
pcvcust_IDRefer=request.form("pccust_IDRefer")
pcvcust_ReferName=request.form("pccust_ReferName")

'MailUp-S
pcIncMailUp=0
query="SELECT pcMailUpSett_RegSuccess FROM pcMailUpSettings WHERE pcMailUpSett_RegSuccess=1;"
set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs = nothing
		call closedb()
		response.Redirect("upddb_MailUp.asp")
	end if
if not rs.eof then
	pcIncMailUp=1
end if
set rs=nothing
'MailUp-E

strSQL="SELECT idcustomer,name,lastName, customerCompany,phone,email,address,address2,city,stateCode,state,zip,CountryCode,customerType,iRewardPointsAccrued,iRewardPointsUsed,RecvNews,IDRefer FROM customers"

set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(strSQL)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<title>Customer Data Export</title>
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

Function GenHTML()%>
	<h1>Customer Data Export</h1>
	<table class="salesExport">
		<tr>
		<%	If idcustomer="1" Then %>
			<th>ID</font></th>
		<% End If
		If pname="1" Then %>
			<th>First Name</font></th>
		<% End If
		If lastName="1" Then %>
			<th>Last Name</font></th>
		<% End If
		if customerCompany="1" Then %>
			<th>Company</font></th>
		<% End If
		if phone="1" Then %>
			<th>Phone</font></th>
		<% End If
		if email="1" Then %>
			<th>Email</font></th>
		<% End If
		if address="1" Then %>
			<th>Address</font></th>
		<% End If
		if address2="1" Then %>
			<th>Address 2</font></th>
		<% End If
		if city="1" Then %>
			<th>City</font></th>
		<% End If
		if stateCode="1" Then %>
			<th>State/Province</font></th>
		<% End If
		if zip="1" Then %>
			<th>Zip</font></th>
		<% End If
		if CountryCode="1" Then %>
			<th>Country</font></th>
		<% End If
		if customerType="1" Then %>
			<th>Customer Type</font></th>
		<% End If 
		if pcvrp_accrued="1" Then %>
			<th><% =RewardsLabel%> Accrued</font></th>
		<% End If 
		if pcvrp_used="1" Then %>
			<th><% =RewardsLabel%> Used</font></th>
		<% End If 
		if pcvrp_available="1" Then %>
			<th>Available <% =RewardsLabel%></font></th>
		<% End If
		For i=1 to fieldCount
		if FieldList(i,0)<>"" Then %>
			<th><%=FieldList(i,1)%></font></th>
		<% End If
		Next
		'MailUp-S
		if pcvcust_recvnews="1" Then %>
			<%if pcIncMailUp=1 then%>
				<th>MailUp Opted-in Lists</th>
			<%else%>
				<th>Newsletter Subscriber</font></th>
			<%end if%>
		<% End If
		'MailUp-E
		if pcvcust_IDRefer="1" Then %>
			<th>Referrer ID</font></th>
		<% End If
		if pcvcust_ReferName="1" Then %>
			<th>Referrer Name</font></th>
		<% End If%>
	</TR>
	<% if (rstemp.BOF=True and rstemp.EOF=True) then%>
		<tr>
			<td>No records found</td>
		</tr>
	<% else
		rstemp.MoveFirst
		Do While Not rstemp.EOF
			pidcustomer=rstemp("idcustomer")
			pcname=rstemp("name")
			plastName=rstemp("lastName")
			pcustomerCompany=rstemp("customerCompany")
			pphone=rstemp("phone")
			pemail=rstemp("email")
			paddress=rstemp("address")
			paddress2=rstemp("address2")
			pcity=rstemp("city")
			pstateCode=rstemp("stateCode")
			pstate=rstemp("state")
			if pstateCode="" then
				pstateCode=pstate
			end if
			pzip=rstemp("zip")
			pCountryCode=rstemp("CountryCode")
			pcustomerType=rstemp("customertype")
			piRewardPointsAccrued=rstemp("iRewardPointsAccrued")
			piRewardPointsUsed=rstemp("iRewardPointsUsed")
			pRecvNews=rstemp("RecvNews")
			pIDRefer=rstemp("IDRefer")
			if pIDRefer<>"" then
			else
				pIDRefer=0
			end if%>
			<TR>
			<%If idcustomer="1" Then %>
				<td><%=pidcustomer%></TD>
			<% End If
			If pname="1" Then %>
			<td><%=pcname%></TD>
			<% End If
			If lastName="1" Then %>
				<td><%=plastName%></TD>
			<% End If
			if customerCompany="1" Then %>
				<td><%=pcustomerCompany%></TD>
			<% End If
			if phone="1" Then %>
				<td><%=pphone%></TD>
			<% End If
			if email="1" Then %>
				<td><%=pemail%></TD>
			<% End If
			if address="1" Then %>
				<td><%=paddress%></TD>
			<% End If
			if address2="1" Then %>
				<td><%=paddress2%></TD>
			<% End If
			if city="1" Then %>
				<td><%=pcity%></TD>
			<% End If
			if stateCode="1" Then %>
				<td><%=pstateCode%></TD>
			<% End If
			if zip="1" Then %>
				<td><%=pzip%></TD>
			<% End If
			if CountryCode="1" Then %>
				<td><%=pCountryCode%></TD>
			<% End If
			if customerType="1" then
				Select Case cint(pcustomerType)
					Case 0: CustomerTypeStr="Retail"
					Case 1: CustomerTypeStr="Wholesale"
				End Select %>
				<td><%=CustomerTypeStr%></TD>
			<% End If
			if pcvrp_accrued="1" Then %>
				<td><%=piRewardPointsAccrued%></TD>
			<% End If
			if pcvrp_used="1" Then %>
				<td><%=piRewardPointsUsed%></TD>
			<% End If
			if pcvrp_available="1" then
				pcvrp_a = (piRewardPointsAccrued-piRewardPointsUsed)%>
				<td><%=pcvrp_a%></TD>
			<% End If
			For i=1 to fieldCount
			if FieldList(i,0)<>"" Then
				query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & pidcustomer & " AND pcCField_ID=" & FieldList(i,0) & ";"
				set rs=connTemp.execute(query)
				fieldValue=""
				if not rs.eof then
					fieldValue=rs("pcCFV_Value")
				end if
				set rs=nothing%>
				<td><%=fieldValue%></TD>
			<% End If
			Next
			'MailUp-S
			if pcvcust_recvnews="1" then
				if pcIncMailUp=0 then
					pcvcust_recvnewsv = pRecvNews%>
					<td><% if pcvcust_recvnewsv = 1 Then Response.Write("Yes") Else Response.Write("No") %></td>
				<%Else 
					query="SELECT pcMailUpLists_ListID FROM pcMailUpLists INNER JOIN pcMailUpSubs ON pcMailUpLists.pcMailUpLists_ID=pcMailUpSubs.pcMailUpLists_ID WHERE idcustomer=" & pidcustomer & " AND pcMailUpSubs_OptOut=0;"
					set rsQ=connTemp.execute(query)
					tmp_MULists=""
					if not rsQ.eof then
						tmpArr=rsQ.getRows()
						intCount=ubound(tmpArr,2)
						For k=0 to intCount
							tmp_MULists=tmp_MULists & tmpArr(0,k) & "|"
						Next
					end if
					set rsQ=nothing%>
					<td><%=tmp_MULists%></td>
				<% End If
			End If
			'MailUp-E
			If pcvcust_IDRefer="1" then%>
				<td><%=pIDRefer%></TD>
			<%End If
			if pcvcust_ReferName="1" then
				if pIDRefer<>"0" then
					query="SELECT name FROM REFERRER where IdRefer="&pIDRefer
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
				end if%>
				<td><%=pcv_ReferName%></TD>
			<%End If%>
		</TR>
		<%rstemp.movenext
		loop
	End if%>
    <tr>
    	<td colspan="50">Report Created On: <%=now()%></td>
</TABLE>

<%
End Function

Function CreateCSVFile()
 	strFile=GenFileName()   
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(server.MapPath(".") & "\" & strFile & ".csv",True)
	If Not rstemp.EOF Then
		set StringBuilderObj = new StringBuilder
		If idcustomer="1" then
			StringBuilderObj.append chr(34) & "Customer ID" & chr(34) & ","
		End If
		If pname="1" then
			StringBuilderObj.append chr(34) & "First Name" & chr(34) & ","
		End If
		If lastName="1" then
			StringBuilderObj.append chr(34) & "Last Name" & chr(34) & ","
		End If
		If customerCompany="1" then
			StringBuilderObj.append chr(34) & "Company" & chr(34) & ","
		End If
		If phone="1" then
			StringBuilderObj.append chr(34) & "Phone" & chr(34) & ","
		End If
		If email="1" then
			StringBuilderObj.append chr(34) & "Email" & chr(34) & ","
		End If
		If address="1" then
			StringBuilderObj.append chr(34) & "Address" & chr(34) & ","
		End If
		If address2="1" then
			StringBuilderObj.append chr(34) & "Address 2" & chr(34) & ","
		End If
		If city="1" then
			StringBuilderObj.append chr(34) & "City" & chr(34) & ","
		End If
		If stateCode="1" then
			StringBuilderObj.append chr(34) & "State/Province" & chr(34) & ","
		End If
		If zip="1" then
			StringBuilderObj.append chr(34) & "Zip" & chr(34) & ","
		End If
		If CountryCode="1" then
			StringBuilderObj.append chr(34) & "Country" & chr(34) & ","
		End If
		If customerType="1" then
			StringBuilderObj.append chr(34) & "Customer Type" & chr(34) & ","
		End If
		If pcvrp_accrued="1" then
			StringBuilderObj.append chr(34) & RewardsLabel & " Accrued" & chr(34) & ","
		End If
		If pcvrp_used="1" then
			StringBuilderObj.append chr(34) & RewardsLabel & " Used" & chr(34) & ","
		End If
		If pcvrp_available="1" then
			StringBuilderObj.append chr(34) & "Available" & RewardsLabel & chr(34) & ","
		End If
		For i=1 to fieldCount
		If FieldList(i,0)<>"" Then
			StringBuilderObj.append chr(34) & FieldList(i,1) & chr(34) & ","
		End If
		Next
		'MailUp-S
		If pcvcust_recvnews="1" then
			If pcIncMailUp=0 then
				StringBuilderObj.append chr(34) & "Newsletter Subscriber" & chr(34) & ","
			Else
				StringBuilderObj.append chr(34) & "MailUp Opted-in Lists" & chr(34) & ","
			End If
		End if
		'MailUp-E
		if pcvcust_IDRefer="1" Then
			StringBuilderObj.append chr(34) & "Referrer ID" & chr(34) & ","
		End If
		if pcvcust_ReferName="1" Then
			StringBuilderObj.append chr(34) & "Referrer Name" & chr(34) & ","
		End If
		a.WriteLine(StringBuilderObj.toString())
		set StringBuilderObj = nothing
		
		Do Until rstemp.EOF
			pidcustomer=rstemp("idcustomer")
			pcname=rstemp("name")
			plastName=rstemp("lastName")
			pcustomerCompany=rstemp("customerCompany")
			pphone=rstemp("phone")
			pemail=rstemp("email")
			paddress=rstemp("address")
			paddress2=rstemp("address2")
			pcity=rstemp("city")
			pstateCode=rstemp("stateCode")
			pstate=rstemp("state")
			if pstateCode="" then
				pstateCode=pstate
			end if
			pzip=rstemp("zip")
			pCountryCode=rstemp("CountryCode")
			pcustomerType=rstemp("customertype")
			piRewardPointsAccrued=rstemp("iRewardPointsAccrued")
			piRewardPointsUsed=rstemp("iRewardPointsUsed")
			pRecvNews=rstemp("RecvNews")
			pIDRefer=rstemp("IDRefer")
			if pIDRefer<>"" then
			else
				pIDRefer=0
			end if
			
			set StringBuilderObj = new StringBuilder
			If idcustomer="1" then 
				StringBuilderObj.append chr(34) & pidcustomer & chr(34) & ","
			End If
			If pname="1" then
				StringBuilderObj.append chr(34) & pcname & chr(34) & ","
			End If
			If lastName="1" then
				StringBuilderObj.append chr(34) & plastName & chr(34) & ","
			End If
			If customerCompany="1" then
				StringBuilderObj.append chr(34) & pcustomerCompany & chr(34) & ","
			End If
			If phone="1" then
				StringBuilderObj.append chr(34) & pphone & chr(34) & ","
			End If
			If email="1" then
				StringBuilderObj.append chr(34) & pemail & chr(34) & ","
			End If
			If address="1" then
				StringBuilderObj.append chr(34) & paddress & chr(34) & ","
			End If
			If address2="1" then
				StringBuilderObj.append chr(34) & paddress2 & chr(34) & ","
			End If
			If city="1" then
				StringBuilderObj.append chr(34) & pcity & chr(34) & ","
			End If
			If stateCode="1" then
				StringBuilderObj.append chr(34) & pstateCode & chr(34) & ","
			End If
			If zip="1" then
				StringBuilderObj.append chr(34) & pzip & chr(34) & ","
			End If
			If CountryCode="1" then
				StringBuilderObj.append chr(34) & pCountryCode & chr(34) & ","
			End If
			If customerType="1" then
				Select Case cint(pcustomerType)
					Case 0: CustomerTypeStr="Retail"
					Case 1: CustomerTypeStr="Wholesale"
				End Select
				StringBuilderObj.append chr(34) & CustomerTypeStr & chr(34) & ","
			End If
			If pcvrp_accrued="1" then
				StringBuilderObj.append chr(34) & piRewardPointsAccrued & chr(34) & ","
			End If
			If pcvrp_used="1" then
				StringBuilderObj.append chr(34) & piRewardPointsUsed & chr(34) & ","
			End If
			If pcvrp_available="1" then
				pcvrp_a = piRewardPointsAccrued-piRewardPointsUsed
				StringBuilderObj.append chr(34) & pcvrp_a & chr(34) & ","
			End If
			For i=1 to fieldCount
			if FieldList(i,0)<>"" Then
				query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & pidcustomer & " AND pcCField_ID=" & FieldList(i,0) & ";"
				set rs=connTemp.execute(query)
				fieldValue=""
				if not rs.eof then
					fieldValue=rs("pcCFV_Value")
				end if
				set rs=nothing
				StringBuilderObj.append chr(34) & fieldValue & chr(34) & ","
			End If
			Next
			'MailUp-S
			If pcvcust_recvnews="1" then
				If pcIncMailUp=0 then
					StringBuilderObj.append chr(34) & pRecvNews & chr(34) & ","
				Else
					query="SELECT pcMailUpLists_ListID FROM pcMailUpLists INNER JOIN pcMailUpSubs ON pcMailUpLists.pcMailUpLists_ID=pcMailUpSubs.pcMailUpLists_ID WHERE idcustomer=" & pidcustomer & " AND pcMailUpSubs_OptOut=0;"
					set rsQ=connTemp.execute(query)
					tmp_MULists=""
					if not rsQ.eof then
						tmpArr=rsQ.getRows()
						intCount=ubound(tmpArr,2)
						For k=0 to intCount
							tmp_MULists=tmp_MULists & tmpArr(0,k) & "|"
						Next
					end if
					set rsQ=nothing
					StringBuilderObj.append chr(34) & tmp_MULists & chr(34) & ","
				End If
			End If
			'MailUp-E
			If pcvcust_IDRefer="1" then
				StringBuilderObj.append chr(34) & pIDRefer & chr(34) & ","
			End If
			if pcvcust_ReferName="1" then
				if pIDRefer<>"0" then
					query="SELECT name FROM REFERRER where IdRefer="&pIDRefer
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
				StringBuilderObj.append chr(34) &pcv_ReferName & chr(34) & ","
			End If
			a.Writeline(StringBuilderObj.toString())
			set StringBuilderObj = nothing
			rstemp.MoveNext			
		Loop
	End If
	a.Close
	Set fs=Nothing
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
	If idcustomer="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Customer ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pname="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="First Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If lastName="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Last Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If customerCompany="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Company"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If phone="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Phone"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If email="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Email"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If address="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Address"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If address2="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Address 2"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If city="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="City"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If stateCode="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="State"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If zip="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Zip"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If CountryCode="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Country"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If customerType="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Customer Type"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pcvrp_accrued="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value=RewardsLabel & " Accrued"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pcvrp_used="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value=RewardsLabel & " Used"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	If pcvrp_available="1" then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Available" & RewardsLabel
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	For i=1 to fieldCount
		If FieldList(i,0)<>"" Then
			t=t+1
			xlWorksheet.Cells(1,t).Value=FieldList(i,1)
			xlWorksheet.Cells(1,t).Interior.ColorIndex=6
		End If
	Next
	'MailUp-S
	If pcvcust_recvnews="1" then
		If pcIncMailUp=0 then
			t=t+1
			xlWorksheet.Cells(1,t).Value="Newsletter Subscriber"
			xlWorksheet.Cells(1,t).Interior.ColorIndex=6
		Else
			t=t+1
			xlWorksheet.Cells(1,t).Value="MailUp Opted-in Lists"
			xlWorksheet.Cells(1,t).Interior.ColorIndex=6
		End If
	End if
	'MailUp-E
	if pcvcust_IDRefer="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Referrer ID"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	if pcvcust_ReferName="1" Then
		t=t+1
		xlWorksheet.Cells(1,t).Value="Referrer Name"
		xlWorksheet.Cells(1,t).Interior.ColorIndex=6
	End If
	iRow=2
	If Not rstemp.EOF Then
		Do Until rstemp.EOF
			pidcustomer=rstemp("idcustomer")
			pcname=rstemp("name")
			plastName=rstemp("lastName")
			pcustomerCompany=rstemp("customerCompany")
			pphone=rstemp("phone")
			pemail=rstemp("email")
			paddress=rstemp("address")
			paddress2=rstemp("address2")
			pcity=rstemp("city")
			pstateCode=rstemp("stateCode")
			pstate=rstemp("state")
			if pstateCode="" then
				pstateCode=pstate
			end if
			pzip=rstemp("zip")
			pCountryCode=rstemp("CountryCode")
			pcustomerType=rstemp("customertype")
			piRewardPointsAccrued=rstemp("iRewardPointsAccrued")
			piRewardPointsUsed=rstemp("iRewardPointsUsed")
			pCI1=rstemp("CI1")
			pCI2=rstemp("CI2")
			pRecvNews=rstemp("RecvNews")
			pIDRefer=rstemp("IDRefer")
			if pIDRefer<>"" then
			else
				pIDRefer=0
			end if
			t=0
			If idcustomer="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pidcustomer
			End If
			If pname="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pcname
			End If
			If lastName="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=plastName
			End If
			If customerCompany="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pcustomerCompany
			End If
			If phone="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pphone
			End If
			If email="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pemail
			End If
			If address="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=paddress
			End If
			If address2="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=paddress2
			End If
			If city="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pcity
			End If
			If stateCode="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pstateCode
			End If
			If zip="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pzip
			End If
			If CountryCode="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pCountryCode
			End If
			If customerType="1" then
				Select Case cint(pcustomerType)
					Case 0: CustomerTypeStr="Retail"
					Case 1: CustomerTypeStr="Wholesale"
				End Select
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=CustomerTypeStr
			End If
			If pcvrp_accrued="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=piRewardPointsAccrued
			End If
			If pcvrp_used="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=piRewardPointsUsed
			End If
			If pcvrp_available="1" then
				pcvrp_a = piRewardPointsAccrued-piRewardPointsUsed
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pcvrp_a
			End If
			For i=1 to fieldCount
			if FieldList(i,0)<>"" Then
				query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & pidcustomer & " AND pcCField_ID=" & FieldList(i,0) & ";"
				set rs=connTemp.execute(query)
				fieldValue=""
				if not rs.eof then
					fieldValue=rs("pcCFV_Value")
				end if
				set rs=nothing
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=fieldValue
			End If
			Next
			'MailUp-S
			If pcvcust_recvnews="1" then
				If pcIncMailUp=0 then
					t=t+1
					xlWorksheet.Cells(iRow,i + t).Value=pRecvNews
				Else
					query="SELECT pcMailUpLists_ListID FROM pcMailUpLists INNER JOIN pcMailUpSubs ON pcMailUpLists.pcMailUpLists_ID=pcMailUpSubs.pcMailUpLists_ID WHERE idcustomer=" & pidcustomer & " AND pcMailUpSubs_OptOut=0;"
					set rsQ=connTemp.execute(query)
					tmp_MULists=""
					if not rsQ.eof then
						tmpArr=rsQ.getRows()
						intCount=ubound(tmpArr,2)
						For k=0 to intCount
							tmp_MULists=tmp_MULists & tmpArr(0,k) & "|"
						Next
					end if
					set rsQ=nothing
					t=t+1
					xlWorksheet.Cells(iRow,i + t).Value=tmp_MULists
				End if
			End If
			'MailUp-E
			If pcvcust_IDRefer="1" then
				t=t+1
				xlWorksheet.Cells(iRow,i + t).Value=pIDRefer
			End If
			if pcvcust_ReferName="1" then
				if pIDRefer<>"0" then
					query="SELECT name FROM REFERRER where IdRefer="&pIDRefer
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
				xlWorksheet.Cells(iRow,i + t).Value=pcv_ReferName
			End If
			iRow=iRow + 1
			rstemp.MoveNext
		Loop
	End If
	strFile=GenFileName()
	xlWorksheet.SaveAs Server.MapPath(".") & "\" & strFile & ".xls"
	xlApplication.Quit
	Set xlWorksheet=Nothing
	Set xlApplication=Nothing
	response.redirect "getFile.asp?file="& strFile &"&Type=xls"
End Function
%>