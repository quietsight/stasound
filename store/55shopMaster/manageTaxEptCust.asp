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
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<% Dim rs, conntemp, query

If request.Form("submitform")="YES" then
	intZoneRateID=getUserInput(request.Form("ZoneRateID"),10)
	strMode=getUserInput(request.Form("mode"),10)
	Main
End If

Sub Main
	Dim intCustomerID, strMsg
	If strMode="view" Then
		call opendb()
		For Each intCustomerID in Request.Form("Customer")
			if isNumeric(intCustomerID) then
				query="DELETE FROM pcTaxEptCust WHERE pcTaxZoneRate_ID="&intZoneRateID&" AND idCustomer=" & intCustomerID
				SET rs=server.CreateObject("ADODB.RecordSet")
				SET rs=conntemp.execute(query)
				strMsg="Customer(s)%20successfully%20removed"
			end if
		Next
		set rs=nothing
		call closedb()
	Else
		call opendb()
		For Each intCustomerID in Request.Form("Customer")
			if isNumeric(intCustomerID) then
				query="INSERT INTO pcTaxEptCust (idCustomer, pcTaxZoneRate_ID) VALUES ("&intCustomerID&","&intZoneRateID&" );"
				SET rs=server.CreateObject("ADODB.RecordSet")
				SET rs=conntemp.execute(query)
				strMsg="Customer(s)%20added%20to%20tax exemption list"
			end if
		Next
		set rs=nothing
		call closedb()
	End If
	Response.Redirect("manageTaxEptCust.asp?ZoneRateID="&intZoneRateID&"&mode=view&msg="&strMsg&"&reqstr="&reqstr)
End Sub

Dim intZoneRateID, strPcCC_Name, strMode
pageTitle="Manage customer exemptions"

if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request("iPageCurrent")
end If

strMode=""

If isNumeric(Request.QueryString("ZoneRateID")) AND Request.QueryString("ZoneRateID") <> "" Then
	intZoneRateID=Request.QueryString("ZoneRateID")
ElseIf isNumeric(Request.QueryString("ZoneRateID")) AND Request.Form("ZoneRateID") <> "" Then
	intZoneRateID=Request.Form("ZoneRateID")
Else
	Response.Redirect "viewTax.asp"
End If


If Request.QueryString("mode") <> "" Then
	strMode=Request.QueryString("mode")
ElseIf Request.Form("mode") <> "" Then
	strMode=Request.Form("mode")
Else
	strMode="view"
End If

Set conntemp=Server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.RecordSet")
call Opendb()

intPcCC_ZoneRateID=intZoneRateID

OrderBy=getUserInput(request("OrderBy"),0)
if isNULL(OrderBy) or OrderBy="" then
	OrderBy="lastName"
end if
SortOrd=getUserInput(request("SortOrd"),0)
if isNULL(SortOrd) or SortOrd="" then
 SortOrd="asc"
end if
%>
<!--#include file="AdminHeader.asp"-->
<form method="POST" action="manageTaxEptCust.asp" name="myForm" class="pcForms">
	<input type="hidden" name="submitform" value="YES">
	<input type="hidden" name="mode" value="<%=strMode%>">
	<input type="hidden" name="OrderBy" value="<%=OrderBy%>">
	<input type="hidden" name="SortOrd" value="<%=SortOrd%>">
	<input type="hidden" name="ZoneRateID" value="<%=intPcCC_ZoneRateID%>">
	<table class="pcCPcontent">
        <tr>
            <td colspan="4" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr> 
            <th width="5%">Select</th>
            <th width="40%" nowrap><a href="manageTaxEptCust.asp?OrderBy=lastName&SortOrd=ASC&ZoneRateID=<%=intZoneRateID%>&mode=<%=request("mode")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="manageTaxEptCust.asp?OrderBy=lastName&SortOrd=Desc&ZoneRateID=<%=intZoneRateID%>&mode=<%=request("mode")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Name</th>
            <th width="30%" nowrap><a href="manageTaxEptCust.asp?OrderBy=customerCompany&SortOrd=ASC&ZoneRateID=<%= intZoneRateID %>&mode=<%=request("mode")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="manageTaxEptCust.asp?OrderBy=CustomerCompany&SortOrd=Desc&ZoneRateID=<%=intZoneRateID%>&mode=<%=request("mode")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Company</th>
            <th width="25%" nowrap><a href="manageTaxEptCust.asp?OrderBy=CustomerType&SortOrd=ASC&ZoneRateID=<%= intZoneRateID %>&mode=<%=request("mode")%>"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="manageTaxEptCust.asp?OrderBy=CustomerType&SortOrd=Desc&ZoneRateID=<%= intZoneRateID %>&mode=<%=request("mode")%>"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a> Type</th>
        </tr>
        <tr>
            <td colspan="4" class="pcCPspacer"></td>
        </tr>
		<% If strMode="edit" Then
            query="SELECT customers.idcustomer, customers.name, customers.lastName, customers.customerCompany, customers.customerType, customers.idCustomerCategory,suspend FROM customers WHERE idcustomer NOT IN (SELECT idcustomer FROM pcTaxEptCust WHERE pcTaxZoneRate_ID="&intZoneRateID&") ORDER BY " & OrderBy & " " & SortOrd
        Else
            query="SELECT pcTaxEptCust.pcTaxEptCust_ID, pcTaxEptCust.idCustomer, pcTaxEptCust.pcTaxZoneRate_ID, customers.name, customers.lastName, customers.customerCompany, customers.customerType FROM pcTaxEptCust INNER JOIN customers ON pcTaxEptCust.idCustomer = customers.idcustomer WHERE (((pcTaxEptCust.pcTaxZoneRate_ID)="&intZoneRateID&")) ORDER BY " & OrderBy & " " & SortOrd
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
            <td colspan="4" align="center"><div class="pcCPmessage">No Customers Found</div></td>
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
                %>
                <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td align="center"> 
                    <input type="checkbox" name="Customer" value="<%= intRSidCustomer %>" class="clearBorder"></td>
                    <td nowrap><a href="modCusta.asp?idcustomer=<%= intRSidCustomer %>"><%= strRSLastName&", "&strRSName  %></a></td>
                    <% if intRSSuspend<>0 then %>
                        &nbsp;<font color="#FF0000">(suspended)</font>
                        <% end if %>
                    <td>
                        <%= strRSCustomerCompany %>
                    </td>
                    <td><p align="center"><%= strRSCustomerType %></p></td>
                </tr>
                <% count=count + 1
                rs.MoveNext
            Loop
            rs.Close
            set rs=nothing
            call closedb()
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
                <td colspan="4"><input type="checkbox" value="ON" onClick="javascript:checkTheBoxes();" class="clearBorder"> Select All Customers</td>
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
                        <a href="manageTaxEptCust.asp?iPageCurrent=<%=iPageCurrent-1%>&OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&ZoneRateID=<%= intZoneRateID %>&mode=<%=request("mode")%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
                    <% end If
                    For I=1 To iPageCount
                        If Cint(I)=Cint(iPageCurrent) Then %>
                            <b><%=I%></b> 
                        <% Else %>
                            <a href="manageTaxEptCust.asp?iPageCurrent=<%=I%>&OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&ZoneRateID=<%= intZoneRateID %>&mode=<%=request("mode")%>"> 
                            <%=I%></a> 
                        <% End If %>
                    <% Next %>
                    <% if CInt(iPageCurrent) < CInt(iPageCount) then %>
                        <a href="manageTaxEptCust.asp?iPageCurrent=<%=iPageCurrent+1%>&OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&ZoneRateID=<%= intZoneRateID %>&mode=<%=request("mode")%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
                    <% end If %>
                </td>
            </tr>
        <% End If %>
        <!-- end page navigation -->
        <% End If %>

	<tr> 
		<td colspan="4"><hr></td>
	</tr>
	<tr> 
		<td colspan="4">
		<% If strMode="edit" Then %>
			<input type="button" value="View Customers" onClick="document.location.href='manageTaxEptCust.asp?OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&ZoneRateID=<%=intZoneRateID%>&mode=view';" class="submit2">
		<% Else %>
			<input type="button" value="Add Customers" onClick="document.location.href='manageTaxEptCust.asp?OrderBy=<%=request("OrderBy")%>&SortOrd=<%=request("SortOrd")%>&ZoneRateID=<%=intZoneRateID%>&mode=edit';" class="submit2">
		<% End If %>
			<input type="button" value="Return to Tax Rule Details" onClick="document.location.href='AddTaxPerZone.asp?idTaxZonesGroupID=<%=intZoneRateID%>&mode=MOD';">
		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->