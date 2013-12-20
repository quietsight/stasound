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
<% 
response.Buffer=true
Response.Expires=0
Dim FieldList(100,2)
dim query, conntemp, rstemp

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

query="SELECT idcustomer,name,lastName, customerCompany,phone,email,address,address2,city,stateCode,state,zip,CountryCode,customerType,iRewardPointsAccrued,iRewardPointsUsed,RecvNews,IDRefer FROM customers"

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
<%response.end
ELSE
		HTMLResult=""
		set StringBuilderObj = new StringBuilder
		If idcustomer="1" then
			StringBuilderObj.append "<td><b>" & "Customer ID"& "</b></td>"
		End If
		If pname="1" then
			StringBuilderObj.append "<td><b>" & "First Name"& "</b></td>"
		End If
		If lastName="1" then
			StringBuilderObj.append "<td><b>" & "Last Name"& "</b></td>"
		End If
		If customerCompany="1" then
			StringBuilderObj.append "<td><b>" & "Company"& "</b></td>"
		End If
		If phone="1" then
			StringBuilderObj.append "<td><b>" & "Phone"& "</b></td>"
		End If
		If email="1" then
			StringBuilderObj.append "<td><b>" & "Email"& "</b></td>"
		End If
		If address="1" then
			StringBuilderObj.append "<td><b>" & "Address"& "</b></td>"
		End If
		If address2="1" then
			StringBuilderObj.append "<td><b>" & "Address 2"& "</b></td>"
		End If
		If city="1" then
			StringBuilderObj.append "<td><b>" & "City"& "</b></td>"
		End If
		If stateCode="1" then
			StringBuilderObj.append "<td><b>" & "State/Province"& "</b></td>"
		End If
		If zip="1" then
			StringBuilderObj.append "<td><b>" & "Zip"& "</b></td>"
		End If
		If CountryCode="1" then
			StringBuilderObj.append "<td><b>" & "Country"& "</b></td>"
		End If
		If customerType="1" then
			StringBuilderObj.append "<td><b>" & "Customer Type"& "</b></td>"
		End If
		If pcvrp_accrued="1" then
			StringBuilderObj.append "<td><b>" & RewardsLabel & " Accrued"& "</b></td>"
		End If
		If pcvrp_used="1" then
			StringBuilderObj.append "<td><b>" & RewardsLabel & " Used"& "</b></td>"
		End If
		If pcvrp_available="1" then
			StringBuilderObj.append "<td><b>" & "Available" & RewardsLabel& "</b></td>"
		End If
		For i=1 to fieldCount
		If FieldList(i,0)<>"" Then
			StringBuilderObj.append "<td><b>" & FieldList(i,1)& "</b></td>"
		End If
		Next
		If pcvcust_recvnews="1" then
			if pcIncMailUp=0 then
				StringBuilderObj.append "<td><b>" & "Newsletter Subscriber"& "</b></td>"
			else
				StringBuilderObj.append "<td><b>" & "MailUp Opted-in Lists"& "</b></td>"
			end if
		End If
		If pcvcust_IDRefer="1" then
			StringBuilderObj.append "<td><b>" & "Referrer ID"& "</b></td>"
		End If
		If pcvcust_ReferName="1" then
			StringBuilderObj.append "<td><b>" & "Refferrer Name"& "</b></td>"
		End If
		HTMLResult="<table><tr>" & StringBuilderObj.toString() & "</tr>"
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
				StringBuilderObj.append "<td>" & pidcustomer& "</td>"
			End If
			If pname="1" then
				StringBuilderObj.append "<td>" & pcname& "</td>"
			End If
			If lastName="1" then
				StringBuilderObj.append "<td>" & plastName& "</td>"
			End If
			If customerCompany="1" then
				StringBuilderObj.append "<td>" & pcustomerCompany& "</td>"
			End If
			If phone="1" then
				StringBuilderObj.append "<td>" & pphone& "</td>"
			End If
			If email="1" then
				StringBuilderObj.append "<td>" & pemail& "</td>"
			End If
			If address="1" then
				StringBuilderObj.append "<td>" & paddress& "</td>"
			End If
			If address2="1" then
				StringBuilderObj.append "<td>" & paddress2& "</td>"
			End If
			If city="1" then
				StringBuilderObj.append "<td>" & pcity& "</td>"
			End If
			If stateCode="1" then
				StringBuilderObj.append "<td>" & pstateCode& "</td>"
			End If
			If zip="1" then
				StringBuilderObj.append "<td>" & pzip& "</td>"
			End If
			If CountryCode="1" then
				StringBuilderObj.append "<td>" & pCountryCode& "</td>"
			End If
			If customerType="1" then
				Select Case cint(pcustomerType)
					Case 0: CustomerTypeStr="Retail"
					Case 1: CustomerTypeStr="Wholesale"
				End Select
				StringBuilderObj.append "<td>" & CustomerTypeStr& "</td>"
			End If
			If pcvrp_accrued="1" then
				StringBuilderObj.append "<td>" & piRewardPointsAccrued& "</td>"
			End If
			If pcvrp_used="1" then
				StringBuilderObj.append "<td>" & piRewardPointsUsed& "</td>"
			End If
			If pcvrp_available="1" then
				pcvrp_a = piRewardPointsAccrued-piRewardPointsUsed
				StringBuilderObj.append "<td>" & pcvrp_a& "</td>"
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
				StringBuilderObj.append "<td>" & fieldValue& "</td>"
			End If
			Next
			If pcvcust_recvnews="1" then
				If pcIncMailUp=0 then
					StringBuilderObj.append "<td>" & pRecvNews& "</td>"
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
					StringBuilderObj.append "<td>" & tmp_MULists & "</td>"
				End If
			End If
			If pcvcust_IDRefer="1" then
				StringBuilderObj.append "<td>" & pIDRefer & "</td>"
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
				StringBuilderObj.append "<td>"  &pcv_ReferName & "</td>"
			End If
			HTMLResult=HTMLResult & "<tr>" & StringBuilderObj.toString() & "</tr>"
			set StringBuilderObj = nothing
			rstemp.MoveNext
		Loop
set rstemp=nothing
HTMLResult=HTMLResult & "</table>"
END IF
closedb()
%>
<% 
Response.ContentType = "application/vnd.ms-excel"
%>
<%=HTMLResult%>