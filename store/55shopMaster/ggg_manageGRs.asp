<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin="1*3*7*"%>
<% 
pageIcon="pcv4_icon_gift.png"
section="layout" 
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<% 
dim query, conntemp, rs, rsTemp, queryFilter, queryOrder

pcv_IdCustomer=request("idcustomer")

if not validNum(pcv_IdCustomer) then
	noCustomer="YES"
	queryFilter=""
	else
	queryFilter=" WHERE pcEv_IDCustomer=" & pcv_IdCustomer
end if

'// Retrieve ordering preference
pcv_strOrder = trim(request("order"))
pcv_strSort = trim(request("sort"))
if pcv_strSort = "" then pcv_strSort = "ASC"
if pcv_strOrder = "" then 
	queryOrder = " ORDER BY pcEv_Date DESC;"
	else
	queryOrder = " ORDER BY " & pcv_strOrder & " " & pcv_strSort & ";"
end if
call openDb()

	query="SELECT pcEv_IDEvent,pcEv_Name,pcEv_Type,pcEv_Date,pcEv_Hide,pcEv_Active,pcEv_IDCustomer FROM pcEvents" & queryFilter & queryOrder
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	
if noCustomer<>"YES" then
	query="SELECT customers.name, customers.lastName FROM customers WHERE idcustomer="&pcv_IdCustomer&";"
	set rsTemp=server.CreateObject("ADODB.RecordSet")
	set rsTemp=conntemp.execute(query)
	pcv_strCustName=rsTemp("name") & " " & rsTemp("lastName")
	set rsTemp=nothing
	pageTitle="Manage Gift Registries for " & pcv_strCustName
	else
	pageTitle="Manage Gift Registries"
end if	
%> 
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<% 
	IF rs.eof THEN
		if noCustomer<>"YES" then
%>
			<tr>
				<td align="center">
					<div class="pcCPmessage">
						This customer has not yet created a gift registry.<br />
						You can <a href="ggg_instGR.asp?idcustomer=<%=pcv_IdCustomer%>">create a new registry</a> on behalf of a customer.
					</div>
				</td>
			</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td align="center">
				<form class="pcForms">
					<input type="button" name="addnew" value="Create New Registry" onclick="javascript:location='ggg_instGR.asp?idcustomer=<%=pcv_IdCustomer%>';" class="submit2">
					&nbsp;
					<input type="button" name="back" value="View/Edit Customer" onclick="location.href='modcusta.asp?idcustomer=<%=pcv_IdCustomer%>';">
				</form>
				</td>
			</tr>
	<% else %>
			<tr>
				<td align="center">
					<div class="pcCPmessage">
						No registries found.<br />
						You can <a href="viewCusta.asp">look for a customer</a>, and then create a registry for him/her.
					</div>
				</td>
			</tr>
<%
		end if
		
	ELSE
%>
<tr>
	<td>
	<table class="pcCPcontent">
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="8">
            <% if not validNum(pcv_IdCustomer) then %>
            To <strong>Add a New Gift Registry</strong>, <a href="viewCusta.asp">look for a customer</a>, view the customer details, and then click on the <strong>View Gift Registries</strong> link. Here is a list of all Gift Registries that have been added to your store (by you or by your customers directly).
            <% else %>
            <a href="ggg_instGR.asp?idcustomer=<%=pcv_IdCustomer%>">Add a new registry for this customer</a>
            &nbsp;|&nbsp;
            <a href="viewCusta.asp">Locate another customer</a>
            <% end if %>
            </td>
		</tr>
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th nowrap><a href="ggg_manageGRs.asp?order=pcEv_Name&sort=ASC"><img src="images/sortasc_blue.gif" alt="Sort Ascending"></a><a href="ggg_manageGRs.asp?order=pcEv_Name&sort=DESC"><img src="images/sortdesc_blue.gif" alt="Sort Descending"></a>&nbsp;Event Name</th>
			<th nowrap><a href="ggg_manageGRs.asp?order=pcEv_Type&sort=ASC"><img src="images/sortasc_blue.gif" alt="Sort Ascending"></a><a href="ggg_manageGRs.asp?order=pcEv_Type&sort=DESC"><img src="images/sortdesc_blue.gif" alt="Sort Descending"></a>&nbsp;Event Type</th>
			<th nowrap><a href="ggg_manageGRs.asp?order=pcEv_Date&sort=ASC"><img src="images/sortasc_blue.gif" alt="Sort Ascending"></a><a href="ggg_manageGRs.asp?order=pcEv_Date&sort=DESC"><img src="images/sortdesc_blue.gif" alt="Sort Descending"></a>&nbsp;Event Date</th>
			<th nowrap>Customer</th>
			<th nowrap>Visibility</th>
			<th nowrap>Status</th>
			<th nowrap colspan="2">Items</th>
		</tr>
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
	<%
	do while not rs.eof
		gIDEvent=rs("pcEv_IDEvent")
		gType=rs("pcEv_Type")
		if gType<>"" then
		else
			gType="N/A"
		end if
		gName=rs("pcEv_Name")
		gDate=rs("pcEv_Date")
		if year(gDate)="1900" then
			gDate=""
		end if
		if gDate<>"" then
			if scDateFrmt="DD/MM/YY" then
				gDate=(day(gDate)&"/"&month(gDate)&"/"&year(gDate))
			else
				gDate=(month(gDate)&"/"&day(gDate)&"/"&year(gDate))
			end if
    end if
		
		pcv_IdCustomer2=rs("pcEv_IDCustomer")
		'Check to see if the customer exists in the DB
		query="SELECT name FROM customers WHERE idcustomer="&pcv_IdCustomer2&";"
		set rsTemp=server.CreateObject("ADODB.RecordSet")
		set rsTemp=conntemp.execute(query)
		if rsTemp.EOF then
			pcv_CustomerCheck="NO"
		end if
		
		if pcv_CustomerCheck<>"NO" then
			query="SELECT customers.name, customers.lastName FROM customers WHERE idcustomer="&pcv_IdCustomer2&";"
			set rsTemp=server.CreateObject("ADODB.RecordSet")
			set rsTemp=conntemp.execute(query)
			pcv_strCustName2=rsTemp("name") & " " & rsTemp("lastName")
			set rsTemp=nothing
		else
			pcv_strCustName2="<font style='color:#FF0000;'>Customer has<br>been deleted</font>"
		end if
		
		gHide=rs("pcEv_Hide")
		if gHide<>"" then
		else
			gHide="0"
		end if
		gActive=rs("pcEv_Active")
		if gActive<>"" then
		else
			gActive="0"
		end if
		query="select sum(pcEP_Qty) as gQty,sum(pcEP_HQty) as gHQty from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_GC=0 group by pcEP_IDEvent"
		set rs1=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs1=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs1.eof then
			gQty=rs1("gQty")
			gHQty=rs1("gHQty")
		else
			gQty="0"
			gHQty="0"
		end if
		set rs1=nothing
		if gQty<>"" then
		else
			gQty="0"
		end if
				
		if gHQty<>"" then
		else
			gHQty="0"
		end if
		%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                <td><a href="ggg_EditGR.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer2%>"><strong><%=gName%></strong></a></td>
                <td><%=gType%></td>
                <td nowrap><%=gDate%></td>
                <td><a href="modCusta.asp?idcustomer=<%=pcv_IdCustomer2%>" target="_blank"><%=pcv_strCustName2%></a></td>
                <td nowrap>
                    <%if gHide="1" then%>
                        Hidden
                    <%else%>
                        Visible
                    <%end if%>
                </td>
                <td nowrap>
                    <%if gActive="1" then%>
                        Active
                    <%else%>
                        <span class="pcCPnotes">Inactive</span>
                    <%end if%>
                </td>
                <td nowrap align="center"><a href="ggg_GRDetails.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer2%>"><%=gQty%></a> (<%=clng(gQty)-clng(gHQty)%>)</td>
                <td nowrap align="center"><a href="ggg_addtoGR.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer2%>"><img src="images/pcIconPlus.jpg" width="12" height="12" alt="Add products to this Gift Registry" title="Add products to this Gift Registry"></a>&nbsp;<a href="ggg_GRDetails.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer2%>"><img src="images/pcIconList.jpg" width="12" height="12" alt="List products" title="List products that have been added to this Gift Registry"></a>&nbsp;<a href="ggg_EditGR.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer2%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit Gift Registry settings" Title="Edit Gift Registry settings"></a>&nbsp;<a href="resultsAdvanced.asp?TypeSearch=registry&pcIntRegistryID=<%=gIDEvent%>&Submit=Search" target="_blank"><img src="images/pcIconSearch.jpg" width="12" height="12" alt="Find orders from this Gift Registry" title="Find orders from this Gift Registry"></a></td>
			</tr>
		<%
		rs.movenext
	loop
	set rs=nothing
	call closeDb()
	%>
	</table>
<%
END IF
%>
	</td>
</tr>
<tr>
    <td colspan="2" class="pcCPspacer">&nbsp;</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->