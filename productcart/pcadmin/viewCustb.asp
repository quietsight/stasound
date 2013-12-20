<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->    
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim rs, conntemp, query

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="lastName"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If

pMode=request.Querystring("mode")
	
call openDb()

Function CreateQuery(Desc,keynum)
Dim m
Dim tmpStr,keywordArray,keylink,keydesc

	tmpStr=""

	Select Case keynum
		Case 1: keydesc="[Name]"
		Case 2: keydesc="LastName"
		Case 3: keydesc="customerCompany"
		Case 4: keydesc="email"
		Case 5: keydesc="city"
		Case 6: keydesc="phone"
		Case 7: keydesc="stateCode"
		Case 8: keydesc="state"
		Case 9: keydesc="zip"
		Case 10: keydesc="CountryCode"
	End Select

	if Instr(Desc," AND ")>0 then
		keywordArray=split(Desc," AND ")
		keylink=" AND "
	else
	if Instr(Desc,",")>0 then
		keywordArray=split(Desc,",")
		keylink=" OR "
	else
		if Instr(Desc," OR ")>0 then
			keywordArray=split(Desc," OR ")
			keylink=" OR "
		else
			keywordArray=split(Desc,"***")
			keylink=" OR "
		end if
	end if
	end if

			
	For m=lbound(keywordArray) to ubound(keywordArray)
	if trim(keywordArray(m))<>"" then
		if tmpStr<>"" then
		tmpStr=tmpStr & keylink
		end if
		tmpStr=tmpStr & "(" & keydesc & " like '%"&trim(keywordArray(m))&"%')"
	end if
	Next
	
	if tmpStr<>"" then
		tmpStr="(" & tmpStr & ")"
	else
		tmpStr="(" & keydesc & " like '%"&Desc&"%')"
	end if

CreateQuery=tmpStr
End Function 

If pMode="ALL" Then
	query="SELECT LastName,[name],customerCompany,phone,customerType,idcustomer,email,idCustomerCategory,pcCust_Locked FROM customers WHERE email<>'REMOVED' ORDER BY "& strORD &" "& strSort
Elseif pMode="LAST" then
	query="SELECT TOP 10 LastName,[name],customerCompany,phone,customerType,idcustomer,email,idCustomerCategory,pcCust_Locked FROM customers WHERE email<>'REMOVED' ORDER BY pcCust_DateCreated DESC;"
Else
	query1=""
	query2=""
	if request("key1")<>"" then
		tmpKey=request("key1")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,1)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key2")<>"" then
		tmpKey=request("key2")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,2)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key3")<>"" then
		tmpKey=request("key3")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,3)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key4")<>"" then
		tmpKey=request("key4")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,4)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key5")<>"" then
		tmpKey=request("key5")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,5)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key6")<>"" then
		tmpKey=request("key6")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,6)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key7")<>"" then
		tmpKey=request("key7")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,7)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key8")<>"" then
		tmpKey=request("key8")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,8)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key9")<>"" then
		tmpKey=request("key9")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,9)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query2=""
	if request("key10")<>"" then
		tmpKey=request("key10")
		tmpKey=replace(tmpKey,"'","''")
		tmpKey=replace(tmpKey,"_","[_]")
		tmpKey=replace(tmpKey,"%","[%]")
		query2=CreateQuery(tmpKey,10)
		if query1<>"" then
			query1=query1 & " AND "
		end if
		if request("key11")="1" then
			query1=query1 & "( NOT (" & query2 & "))"
		else
			query1=query1 & query2
		end if
	end if
	
	query2=""
	if request("customerType")<>"" then
		tmpKey=request("customerType")
		Select Case tmpKey
			Case "0": query2="customerType=0"
			Case "1": query2="customerType=1"
			Case Else: 
			if instr(tmpKey,"CC_")>0 then
				tmpA=split(tmpKey,"CC_")
				query2="idCustomerCategory=" & tmpA(1)
			else
				query2="customerType=0"
			end if
		End Select
		if query1<>"" then
			query1=query1 & " AND "
		end if
		query1=query1 & query2
	end if
	
	query="SELECT LastName,[name],customerCompany,phone,customerType,idcustomer,email,idCustomerCategory,pcCust_Locked FROM customers "
	if query1<>"" then
		query=query & " WHERE " & query1
	end if
	query=query&" ORDER BY "& strORD &" "& strSort &", idCustomerCategory ASC;"
End If

Set rs=Server.CreateObject("ADODB.Recordset")
	
	rs.CacheSize=15
	rs.PageSize=15
	
	rs.Open query, connTemp, adOpenStatic, adLockReadOnly
	
if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "msg.asp?message=26"
end if

	pcv_totCustomers = rs.RecordCount
	rs.MoveFirst

	' get the max number of pages
	Dim iPageCount
	iPageCount=rs.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
	
	' set the absolute page
	rs.AbsolutePage=iPageCurrent

pageTitle="Customer Search Results" 
pageIcon="pcv4_icon_people.png"
pcStrPageName="viewCustb.asp"
section="mngAcc" 
%>
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="" class="pcForms">
<table class="pcCPcontent">

<tr>
	<td><% if pcv_totCustomers <> "" and pcv_totCustomers > 0 then %>Your search returned <%=pcv_totCustomers%>&nbsp;customer(s). <% end if %>The <img src="images/pcadmin_lockedaccount.jpg"> next to a customer's name indicates that the customer's account has been locked by the store administrator (customer can no longer log into the store). To unlock the account, &quot;Edit&quot; the customer and uncheck that option.</td>
	</tr>
</table>
<br>
<table class="pcCPcontent">
<tr>
<th nowrap><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=lastName&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=lastName&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Name</th>
<th nowrap><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=phone&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=phone&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Phone</th>
<th nowrap><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=customerCompany&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=customerCompany&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Company</th>
<th nowrap><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=customerType&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent%>&order=customerType&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Type</b></th>
<th align="right"><a href="viewCusta.asp">New Search</a></th>
</tr>
<tr>
	<td colspan="5" class="pcCPspacer"></td>
</tr>
<% Dim strCol
strCol="#E1E1E1"
Count=0
Do While NOT rs.EOF And Count < rs.PageSize
	pLastName=rs("LastName")
	pname=rs("name")
	pcustomerCompany=rs("customerCompany")
	pphone=rs("phone")
	pcustomerType=rs("customerType")
	pidcustomer=rs("idcustomer")
	pemail=rs("email")
	pcv_IDCustomerCategory=rs("idCustomerCategory")
	if IsNull(pcv_IDCustomerCategory) or (pcv_IDCustomerCategory="") then
		pcv_IDCustomerCategory=0
	end if
	pcCust_Locked=rs("pcCust_Locked")
%>
<tr valign="top" onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
	<td>
		<a href="modCusta.asp?idcustomer=<%=pidcustomer%>"><%=(pLastName&", "&pname)%></a>
		<% if pcustomerType="3" or pcCust_Locked="1" then%>
		<img src="images/pcadmin_lockedaccount.jpg" width="12" height="12"> 
		<% end if %>
	</td>
	<td nowrap><%=pphone%></td>
	<td><a href="modCusta.asp?idcustomer=<%=pidcustomer%>"><%=pcustomerCompany%></a></td>
	<td nowrap>  
		<% if pcustomerType="1" then
			custType="Wholesale"
		else
			custType="Retail"
		end if %>
		<%=custType%>
        <span class="pcSmallText">
        <% if pcf_GetCustType(pidcustomer)>0 then%> - Guest account<% end if %>
        </span>
		<%pcv_CatName=""
		if pcv_IDCustomerCategory<>"0" then
			query="SELECT pcCC_Name FROM pcCustomerCategories WHERE idCustomerCategory=" & pcv_IDCustomerCategory
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcv_CatName=rsQ("pcCC_Name")
			end if
			set rsQ=nothing
		end if
		if pcv_CatName<>"" then
			response.write "<br>" & pcv_CatName
		end if
		%>
    </td>
	<td nowrap align="right" class="cpLinksList"> 
		<a href="mailto:<%=pemail%>">E-Mail</a> | <a href="modCusta.asp?idcustomer=<%=pidcustomer%>">Edit</a> | <a href="viewCustOrders.asp?idcustomer=<%=pidcustomer%>">Orders</a> | <% if pcf_GetCustType(pidcustomer)=0 then%><a href="adminPlaceOrder.asp?idcustomer=<%=pidcustomer%>" target="_blank">Place Order</a> | <%end if%><a href="javascript:if (confirm('Please note: the following will occur if you click on OK. If there ARE NO orders associated with this account, the customer account will be permanently deleted from the database. If the customer had initiated, but not completed one or more orders (i.e. incomplete orders), they will also be deleted. If there ARE orders associated with this account, you will be prompted to either Remove the customer account, but keep the associated orders in the database, OR Delete the customer account and all of the associated orders. Are you sure to want to continue?')) location='delCustomer.asp?idcustomer=<%=pidcustomer%>'">Remove</a>
	</td>
</tr>
                      
<% count=count + 1
rs.MoveNext
loop
%>
</table>
<br>
<table class="pcCPcontent">
<tr>
	<td>
		<% If iPageCount>1 Then %>
			<%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & "<br>")%>
			<p class="pcPageNav">
				<%if iPageCurrent > 1 then %>
					<a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&key7=<%=request("key7")%>&key8=<%=request("key8")%>&key9=<%=request("key9")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a>
				<% end If
				For I=1 To iPageCount
					If Cint(I)=Cint(iPageCurrent) Then %>
						<%=I%> 
					<% Else %>
						<a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&key7=<%=request("key7")%>&key8=<%=request("key8")%>&key9=<%=request("key9")%>&mode=<%=pMode%>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
					<% End If %>
				<% Next %>
				<% if CInt(iPageCurrent) < CInt(iPageCount) then %>
					<a href="viewCustb.asp?customerType=<%=request("customerType")%>&key1=<%=request("key1")%>&key2=<%=request("key2")%>&key3=<%=request("key3")%>&key4=<%=request("key4")%>&key5=<%=request("key5")%>&key6=<%=request("key6")%>&key7=<%=request("key7")%>&key8=<%=request("key8")%>&key9=<%=request("key9")%>&mode=<%=pMode%>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a>
				<% end If %>
			</p>
		<% End If %>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="button" name="Button" value="Back" onClick="javascript:history.back()">&nbsp;
		<input type="button" value="View All" onClick="location.href='viewCustb.asp?mode=ALL'">&nbsp;
		<input type="button" value="Add New" onClick="location.href='instCusta.asp'">&nbsp;
		<input type="button" value="New Search" onClick="location.href='viewCusta.asp'">
	</td>
</tr>
</table>
</form>
<% 
set rs = nothing
call closeDb()
%><!--#include file="AdminFooter.asp"-->