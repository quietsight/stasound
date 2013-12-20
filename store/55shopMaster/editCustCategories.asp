<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<% Dim rs, conntemp, query

If request("submitform")="YES" then
	intCCID=getUserInput(request("CCID"),10)
	strMode=getUserInput(request("mode"),10)
	intWholesalePriv=getUserInput(request("WholesalePriv"),10)
	Main
End If

Sub Main
	Dim intCustomerID, strMsg
	If strMode="view" Then
		call opendb()
		For Each intCustomerID in request("Customer")
			if isNumeric(intCustomerID) then
				query="UPDATE customers SET customerType=0, idCustomerCategory=0 WHERE idCustomer=" & intCustomerID
				SET rs=server.CreateObject("ADODB.RecordSet")
				SET rs=conntemp.execute(query)
				strMsg="Customer(s)%20Removed%20From%20Customer Category Type"
			end if
		Next
		call closedb()
	Else
		call opendb()
		tmpArr=split(request("custlist"),",")
		For i=lbound(tmpArr) to ubound(tmpArr)
			intCustomerID=trim(tmpArr(i))
			if intCustomerID<>"" then
				if isNumeric(intCustomerID) then
					query="UPDATE customers SET customerType="&intWholesalePriv&",idCustomerCategory="&intCCID&" WHERE idCustomer=" & intCustomerID
					SET rs=server.CreateObject("ADODB.RecordSet")
					SET rs=conntemp.execute(query)
					strMsg="Customer(s)%20Added%20to%20Customer Category Type"
				end if
			end if
		Next
		call closedb()
	End If
	Response.Redirect("editCustCategories.asp?CCID="&intCCID&"&mode=view&msg="&strMsg&"&reqstr="&reqstr)
End Sub

Dim intCCID, strPcCC_Name, strMode

if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request("iPageCurrent")
end If

strMode=""
	
If isNumeric(Request.QueryString("CCID")) AND Request.QueryString("CCID") <> "" Then
	intCCID=Request.QueryString("CCID")
ElseIf isNumeric(Request.QueryString("CCID")) AND request("CCID") <> "" Then
	intCCID=request("CCID")
Else
	Response.Redirect "AdmincustomerCategory.asp"
End If

If Request.QueryString("mode") <> "" Then
	strMode=Request.QueryString("mode")
ElseIf request("mode") <> "" Then
	strMode=request("mode")
Else
	strMode="view"
End If

Set conntemp=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.Recordset")
call Opendb()
query="SELECT pcCustomerCategories.idCustomerCategory, pcCustomerCategories.pcCC_Name, pcCustomerCategories.pcCC_Description, pcCustomerCategories.pcCC_WholesalePriv FROM pcCustomerCategories WHERE idCustomerCategory=" & intCCID & " ORDER BY pcCC_Name"
rs.Open query, conntemp, adOpenStatic, adLockReadOnly

strPcCC_Name=rs("pcCC_Name")
intPcCC_WholesalePriv=rs("pcCC_WholesalePriv")
rs.Close

OrderBy=getUserInput(request("OrderBy"),0)
if isNULL(OrderBy) or OrderBy="" then
	OrderBy="lastName"
end if
SortOrd=getUserInput(request("SortOrd"),0)
if isNULL(SortOrd) or SortOrd="" then
 SortOrd="asc"
end if

If strMode="edit" Then
	pageTitle="Add Customers To "
	else
	pageTitle="Current Customers of " 
end if
pageTitle=pageTitle & "&quot;" & strPcCC_Name & "&quot;"
pcStrPageName="editCustCategories.asp"
%>
<!--#include file="AdminHeader.asp"-->
<%
IF strMode="edit" THEN
call opendb()
%>
	<table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr> 
			<td width="100%">
				<table id="FindProducts" class="pcCPcontent">
                    <tr>
                        <td>
                        <%
                        src_FormTitle1="Find Customers"
                        src_FormTitle2="Add New Customer(s) to &quot;" & strPcCC_Name & "&quot;"
                        src_FormTips1="Use the following filters to look for customers in your store."
                        src_FormTips2="Select one or more customers that you want to add to &quot;" & strPcCC_Name & "&quot;"
                        src_DisplayType=1
                        src_ShowLinks=0
                        src_FromPage="editCustCategories.asp?OrderBy=" & request("OrderBy") & "&SortOrd=" & request("SortOrd") & "&CCID=" & intCCID & "&WholesalePriv=" & intPcCC_WholesalePriv & "&mode=edit"
                        src_ToPage="editCustCategories.asp?OrderBy=" & request("OrderBy") & "&SortOrd=" & request("SortOrd") & "&CCID=" & intCCID & "&WholesalePriv=" & intPcCC_WholesalePriv & "&mode=edit&submitform=YES"
                        src_Button1=" Search "
                        src_Button2=" Add customer(s) to '" & strPcCC_Name & "'"
                        src_Button3=" Back "
                        src_PageSize=15
                        UseSpecial=1
                        session("srcCust_from")=""
                        session("srcCust_where")=" customers.idCustomerCategory=0 "
                        %>
                        <!--#include file="inc_srcCusts.asp"-->
                        </td>
                    </tr>
				</table>
			</td>
		</tr>
	</table>
<%
call closedb()
ELSE
%>
<form method="POST" action="editCustCategories.asp" name="myForm" class="pcForms">
	<input type="hidden" name="submitform" value="YES">
	<input type="hidden" name="mode" value="<%=strMode%>">
	<input type="hidden" name="CCID" value="<%= intCCID %>">
	<input type="hidden" name="OrderBy" value="<%=OrderBy%>">
	<input type="hidden" name="SortOrd" value="<%=SortOrd%>">
	<input type="hidden" name="WholesalePriv" value="<%=intPcCC_WholesalePriv%>">
	
    <table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr> 
			<td width="100%"> 
				<table class="pcCPcontent">
					<tr>
						<th width="5%">Select</th>
						<th width="35%"><a href="editCustCategories.asp?OrderBy=lastName&SortOrd=ASC&CCID=<%=intCCID%>&mode=<%=request("mode")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="editCustCategories.asp?OrderBy=lastName&SortOrd=Desc&CCID=<%=intCCID%>&mode=<%=request("mode")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Name</th>
						<th width="35%"><a href="editCustCategories.asp?OrderBy=customerCompany&SortOrd=ASC&CCID=<%= intCCID %>&mode=<%=request("mode")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="editCustCategories.asp?OrderBy=CustomerCompany&SortOrd=Desc&CCID=<%=intCCID%>&mode=<%=request("mode")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Company</th>
						<th width="25%" nowrap><a href="editCustCategories.asp?OrderBy=CustomerType&SortOrd=ASC&CCID=<%= intCCID %>&mode=<%=request("mode")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="editCustCategories.asp?OrderBy=CustomerType&SortOrd=Desc&CCID=<%= intCCID %>&mode=<%=request("mode")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Type</th>
					</tr>
					<tr>
						<td colspan="4" class="pcCPspacer"></td>
					</tr>
					<% If strMode="edit" Then
						query="SELECT customers.idcustomer, customers.name, customers.lastName, customers.customerCompany, customers.customerType, customers.idCustomerCategory,suspend FROM customers WHERE customers.idCustomerCategory=0 ORDER BY " & OrderBy & " " & SortOrd
					Else
						query="SELECT customers.idcustomer, customers.name, customers.lastName, customers.customerCompany, customers.customerType, customers.idCustomerCategory,suspend FROM customers WHERE idCustomerCategory=" & intCCID & " ORDER BY " & OrderBy & " " & SortOrd
					End If
					set rs=Server.CreateObject("ADODB.Recordset")
					rs.CacheSize=50
					rs.PageSize=50

					rs.Open query, conntemp, adOpenStatic, adLockReadOnly
					If Err Then
						rs.Close
						conntemp.Close
					ElseIf rs.EOF OR rs.BOF Then
					%>
					<tr> 
						<td colspan="4" align="center">No Customers Found</td>
					</tr>
				<% Else
						rs.MoveFirst
						' get the max number of pages
						Dim iPageCount
						iPageCount=rs.PageCount
						If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
						If iPageCurrent < 1 Then iPageCurrent=1
						
						' set the absolute page
						rs.AbsolutePage=iPageCurrent
						'end for nav

					Dim Count
					Count=0
					Do While NOT rs.EOF And Count < rs.PageSize
						intRSIdCustomer=rs("idcustomer")
						strRSName=rs("name")
						strRSLastName=rs("lastName")
						strRSCustomerCompany=rs("customerCompany")
						intRSCustomerType=rs("customerType")
						if intRSCustomerType=0 then
							strRSCustomerType="Retail"
						else
							strRSCustomerType="Wholesale"
						end if
						intRSidCustomerCategory=rs("idCustomerCategory")
						intRSSuspend=rs("suspend")
						%>
						<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
							<td align="center"><input type="checkbox" name="Customer" value="<%= intRSidCustomer %>" class="clearBorder"></td>
							<td>
                            <a href="modCusta.asp?idcustomer=<%= intRSidCustomer %>"><%= strRSLastName&", "&strRSName  %></a>
							<% if intRSSuspend<>0 then %>&nbsp;<span class="pcCPnotes">(suspended)</span><% end if %>
                            </td>
							<td><%= strRSCustomerCompany %></td>
							<td><%= strRSCustomerType %></td>
						</tr>
						<% 
						count=count + 1
						rs.MoveNext
					Loop
					%>
					<script>
					function checkTheBoxes()
					{
						with (document.myForm) {
							for(i=0; i < elements.length -1; i++) {
								if ( elements[i].name== "Customer" ) {
									elements[i].checked=true;
								}
							}
						}
					}
					</script>	
					<tr> 
						<td width="36" align="center"> 
						<input type="checkbox" value="ON" onClick="javascript:checkTheBoxes();" class="clearBorder">
						</td>
						<td colspan="3"><b>Select All Customers</b></td>
					</tr>
					<tr>
						<td colspan="4" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="4">
						<input type="submit" value="<%If strMode="edit" Then %>Add <% Else %>Remove <% End If %>Checked" class="submit2">
						</td>
					</tr>
					
					<!-- page navigation-->
					<% If iPageCount>1 Then %>
					<tr>
						<td colspan="4">
						<%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & " <P>")%>
						</td>
					</tr>
					<tr>
						<td colspan="4">
							<%' display Next / Prev buttons
							if iPageCurrent > 1 then %>
								<a href="editCustCategories.asp?iPageCurrent=<%=iPageCurrent-1%>&OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&CCID=<%= intCCID %>&mode=<%=request("mode")%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
							<% end If
							For I=1 To iPageCount
								If Cint(I)=Cint(iPageCurrent) Then %>
									<b><%=I%></b> 
								<% Else %>
									<a href="editCustCategories.asp?iPageCurrent=<%=I%>&OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&CCID=<%= intCCID %>&mode=<%=request("mode")%>"> 
									<%=I%></a> 
								<% End If %>
							<% Next %>
							<% if CInt(iPageCurrent) < CInt(iPageCount) then %>
								<a href="editCustCategories.asp?iPageCurrent=<%=iPageCurrent+1%>&OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&CCID=<%= intCCID %>&mode=<%=request("mode")%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
							<% end If %>
						</td>
					</tr>
				<% End If %>
				<!-- end page navigation -->
				<% End If %>
			</table>
		</td>
	</tr>
	<tr> 
		<td align="center">
		<% If strMode="edit" Then %>
			<input type="button" value="View Customers of <% Response.Write strPcCC_Name %>" onClick="document.location.href='editCustCategories.asp?OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&CCID=<%=intCCID%>&mode=view';" class="submit2">
		<% Else %>
			<input type="button" value="Add Customers to <% Response.Write strPcCC_Name %>" onClick="document.location.href='editCustCategories.asp?OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&CCID=<%=intCCID%>&mode=edit';" class="submit2">
		<% End If %>
		</td>
	</tr>
	<tr>
		<td align="center">
		<input type="button" value="Edit Customer Category Type" onClick="document.location.href='AdminCustomerCategory.asp?mode=2&id=<%=intCCID%>'" name="button">
		&nbsp;
		<input type="button" value="Manage Customer Category Types" onClick="document.location.href='AdminCustomerCategory.asp';" name="button">
		</td>
	</tr>
</table>
</form>
<%
Set rs=Nothing
call closeDb()
END IF
%>
<!--#include file="AdminFooter.asp"-->