<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/statusAPP.inc"-->
<!--#include file="../includes/ErrorHandler.asp"-->

<% 
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If 

dim query, conntemp, rstemp, rs, iPageCurrent

Const iPageSize=10

if request.querystring("iPageCurrent")="" or request.querystring("iPageCurrent")="0" then
	iPageCurrent=1 
else
	iPageCurrent=Request.QueryString("iPageCurrent")
end if

call openDb()

If statusAPP="1" Then
	query="SELECT Distinct orders.idOrder, orders.orderDate, orders.ord_OrderName, orders.OrderStatus FROM pcDropShippersSuppliers,Products,ProductsOrdered INNER JOIN orders ON ProductsOrdered.idOrder = orders.idOrder WHERE ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND products.idproduct=ProductsOrdered.idproduct  AND ((pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct) OR (pcDropShippersSuppliers.idproduct=products.pcprod_ParentPrd)) AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & " AND ((orders.orderStatus>2 AND orders.orderStatus<5) 		OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) ORDER BY orders.idOrder DESC;"
Else
	query="SELECT DISTINCT orders.idOrder, orders.orderDate, orders.ord_OrderName, orders.OrderStatus FROM pcDropShippersSuppliers INNER JOIN (orders INNER JOIN ProductsOrdered ON orders.idorder = ProductsOrdered.idorder) ON (pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper = " & session("pc_sdsIsDropShipper") & ")  WHERE ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) ORDER BY orders.idOrder DESC;"
end if
set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CursorLocation=adUseClient
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, conntemp

if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=34"     
else
	rstemp.MoveFirst
	' get the max number of pages
	Dim iPageCount
	iPageCount=rstemp.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
	' set the absolute page
	rstemp.AbsolutePage=iPageCurrent
end if          

%> 

<!--#include file="header.asp"-->
<div id="pcMain">
	<table class="pcMainTable">   
		<tr>
			<td>
				<h1><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_4")%></h1>
			</td>
		</tr>
		<tr>
			<td>     
			<table class="pcShowContent">
				<tr>
			    <th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_5")%></th>
					<%if scOrderNumber="1" then 'Show order name %>
						<th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_9")%></th>
					<% end if %>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_6")%></th>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_sds_viewpast_1a")%></th>
					<th>&nbsp;</th>
				</tr>
				<tr class="pcSpacer">
					<td colspan="5"></td>
				</tr>
				<%
					Dim mcount
					mcount=0
					Do while (not rstemp.eof) and (mcount<rstemp.PageSize)
						mcount=mcount+1
						pIdOrder = rstemp("idOrder")
						pOrderName = rstemp("ord_OrderName")
						pOrderDate = rstemp("orderDate")
						pOrderStatus=rstemp("OrderStatus")
						
						'// Check to see if pcDropShippersOrders contains a Drop Shipper order status
						query="SELECT pcDropShipO_OrderStatus FROM pcDropShippersOrders WHERE pcDropShipO_idOrder=" & pIdOrder & " AND pcDropShipO_DropShipper_ID=" & session("pc_idsds")
						set rs=Server.CreateObject("ADODB.Recordset")
						set rs=conntemp.execute(query)
						
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						
						if not rs.eof then
							pOrderStatus=rs("pcDropShipO_OrderStatus")
						end if     
						
						if IsNull(pOrderStatus) or pOrderStatus="" then
							pOrderStatus=0
						end if
				%>
				<tr>
					<td>
						<a href="sds_viewPastD.asp?idOrder=<%response.write (scpre+int(pIdOrder))%>"><%response.write (scpre+int(pIdOrder))%></a>
					</td>
					<%if scOrderNumber="1" then 'Show order name %>
					<td>
						<%=pOrderName%>
					</td>
					<% end if %>
					<td>
						<%=showdateFrmt(pOrderDate)%>
					</td>
					<td nowrap>
						<%Select Case pOrderStatus
						Case 2: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_2")
						Case 3: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_3")
						Case 4: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")
						Case 5: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_5")
						Case 6: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_6")
						Case 9: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_9")
						Case 10: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")
						Case 12: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")						
						Case Else:
							queryQ="SELECT idProductOrdered FROM ProductsOrdered WHERE idOrder=" & pIdOrder & " AND pcDropShipper_ID=" & session("pc_idsds") & " AND pcPrdOrd_Shipped=0;"
							set rsQ=connTemp.execute(queryQ)
							if not rsQ.eof then
								queryQ="SELECT idProductOrdered FROM ProductsOrdered WHERE idOrder=" & pIdOrder & " AND pcDropShipper_ID=" & session("pc_idsds") & " AND pcPrdOrd_Shipped=1;"
								set rsQ=connTemp.execute(queryQ)
								if not rsQ.eof then
									response.write dictLanguage.Item(Session("language")&"_sds_viewpast_7")
								else
									response.write dictLanguage.Item(Session("language")&"_sds_viewpast_3")
								end if
								set rsQ=nothing
							else
								response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")
							end if
							set rsQ=nothing
						End Select%>
					</td>
					<td nowrap>
						<div align="right" class="pcSmallText">
							<a href="sds_viewPastD.asp?idOrder=<%response.write (scpre+int(pIdOrder))%>"><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_3")%></a><%if pOrderStatus="3" or pOrderStatus="7" or pOrderStatus="8" then%> - <a href="sds_ShipOrderWizard1.asp?idOrder=<%=pIdOrder%>"><%response.write dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%></a><%end if%>
						</div>
					</td>
				</tr>
				<%
				rstemp.movenext
			  loop
				%>
			</table>
  			<% 
			set rstemp = nothing
			call closeDb()
			%>
			</td>
		</tr>
		<tr>
			<td><hr></td>
		</tr>
        
		<% 
        if iPageCount>1 Then 
        %>

			<td> 
				<%Response.Write("Page "& iPageCurrent & " of "& iPageCount & "<br />")%>
				<%'Display Next / Prev buttons
				if iPageCurrent > 1 then
					'We are not at the beginning, show the prev button %>
					<a href="sds_viewPast.asp?iPageCurrent=<%=iPageCurrent-1%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
				<% end If
				If iPageCount <> 1 then
					For I=1 To iPageCount
						If int(I)=int(iPageCurrent) Then %>
							<%=I%> 
						<% Else %>
							<a href="sds_viewPast.asp?iPageCurrent=<%=I%>" style="text-decoration: underline;"><%=I%></a> 
						<% End If %>
					<% Next %>
				<% end if %>
				<% if CInt(iPageCurrent) <> CInt(iPageCount) then
					'We are not at the end, show a next link %>
					<a href="sds_viewPast.asp?iPageCurrent=<%=iPageCurrent+1%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
				<% end If %>
			</td>
		</tr>
	<tr>
		<td><hr></td>
	</tr>          
	</table>
<% end if %>

        
        
		<tr> 
			<td><a href="sds_MainMenu.asp"><img src="<%=rslayout("back")%>"></a></td>
		</tr>
	</table>
</div>
<!--#include file="footer.asp"-->