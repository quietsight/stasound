<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Dim pageTitle, pageName, pageIcon, Section
pageTitle="Subscriptions Report"
pageName="sb_ViewSubs.asp"
pageIcon="pcv4_icon_sb.png"
Section="SB" 
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="sb_inc.asp"-->
<% dim conntemp, query, rs

call opendb()

pcv_intIDMain = request("idmain")

Const iPageSize=10
Dim iPageCurrent
if request.querystring("iPageCurrent")="" or request.querystring("iPageCurrent")="0" then
	iPageCurrent=1
else
	iPageCurrent=Request.QueryString("iPageCurrent")
end if

If pcv_intIDMain="0" Then
	query="SELECT orders.idOrder, orders.orderDate, orders.total, orders.ord_OrderName, ProductsOrdered.idProductOrdered,ProductsOrdered.UnitPrice,ProductsOrdered.quantity, ProductsOrdered.pcSubscription_ID, ProductsOrdered.pcPO_SubAmount, ProductsOrdered.pcPO_SubActive, ProductsOrdered.pcPO_IsTrial, ProductsOrdered.pcPO_SubTrialAmount, ProductsOrdered.pcPO_SubStartDate, ProductsOrdered.pcPO_SubType, productsordered.pcPO_LinkID FROM orders, productsordered WHERE orders.OrderStatus>1  And orders.idOrder = ProductsOrdered.idOrder and ProductsOrdered.pcSubscription_ID >0  ORDER BY ProductsOrdered.idProductOrdered DESC"
Else
	query="SELECT orders.idOrder, orders.orderDate, orders.total, orders.ord_OrderName, ProductsOrdered.idProductOrdered,ProductsOrdered.UnitPrice,ProductsOrdered.quantity, ProductsOrdered.pcSubscription_ID, ProductsOrdered.pcPO_SubAmount, ProductsOrdered.pcPO_SubActive, ProductsOrdered.pcPO_IsTrial, ProductsOrdered.pcPO_SubTrialAmount, ProductsOrdered.pcPO_SubStartDate, ProductsOrdered.pcPO_SubType, productsordered.pcPO_LinkID FROM orders, productsordered WHERE productsordered.pcPO_LinkID='" & pcv_intIDMain &"' AND orders.OrderStatus>1  And orders.idOrder = ProductsOrdered.idOrder and ProductsOrdered.pcSubscription_ID >0  ORDER BY ProductsOrdered.idProductOrdered DESC"
End If

set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.CursorLocation=adUseClient
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, conntemp

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error in sb_CustViewSubs.asp: "&err.description) 
end If
if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=34"     
else
	rstemp.MoveFirst
	rstemp.AbsolutePage=iPageCurrent
end if          

' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<form method="POST" action="" name="checkboxform" class="pcForms">
	<table class="pcCPcontent">   
		<tr>
			<td>
				<h1>
					<%
					If pcv_intIDMain="0" Then
						response.Write("View All Subscription Orders")
					Else
						response.Write("View Subscription Orders")
					End If 					
					%>
                </h1>
			</td>
		</tr>
		<tr>
			<td>     
            
			<%
            '// Get the max number of pages
            Dim iPageCount
            iPageCount=rstemp.PageCount				
            session("CP_OrdSrcPages")= iPageCount
            If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
            If iPageCurrent < 1 Then iPageCurrent=1
                
            if rstemp.eof then 
                presults="0"
            else
                %>
                <table class="pcCPcontent">
                    <tr> 
                        <td>
                            <% ' Showing total number of pages found and the current page number %>
                            Displaying Page <b><%=iPageCurrent%></b> of <b><%=iPageCount%></b>
                            &nbsp;|&nbsp;
                            Total Records Found: <b><%=rstemp.RecordCount%></b>
                        </td>
                    </tr>
                </table>
                <% 
            end if '// if rstemp.eof then
            %>  
            
			<table class="pcShowContent">
				<tr>
			    	<th nowrap>Order ID</th>
                    <th nowrap>Order Date</th>
                    <th nowrap>Package ID</th>
                	<th nowrap>Subscription ID</th> 
					<th nowrap>Price</th>
					<th nowrap>Trial Price</th>                    
					<th></th>
				</tr>
				<tr class="pcSpacer">
					<td colspan="5"></td>
				</tr>
				<%
				'// Showing relevant records
				Dim rcount, i, x
	
				Do while (not rstemp.eof) and (mcount<rstemp.PageSize)
				
				'do while not rstemp.eof
					idorder = rstemp("idOrder")
					idProductOrdered = rstemp("idProductOrdered")
					pSubUnitPrice = rstemp("unitPrice")
					pSubQty = rstemp("quantity")
					pSubPrice = rstemp("pcPO_SubAmount")
					pSubTrial = rstemp("pcPO_IsTrial")
					pSubTrialAmount = rstemp("pcPO_SubTrialAmount")						 
					pSubStartDate = rstemp("pcPO_SubStartDate")
					pSubActive =rstemp("pcPO_SubActive") 
					pSubType=rstemp("pcPO_SubType")	
					pLinkID=rstemp("pcPO_LinkID")

					'// Obtain Status
					Dim pvc_Status
					if pSubActive = "1"  Then
						pcv_Status = "Active"
					Elseif pSubActive = "2" then
						pcv_Status = "Pending"
					else
						pcv_Status = "<font color='#ff0000'>Not-Active</font>"
					End if 
					
					mcount=mcount+1
					
					'// Obtain GUID and Email
					Dim pcv_strCustEmail, pcv_strGUID
					pcv_strCustEmail=""
					pcv_strGUID=""
					query = "SELECT customers.email, SB_Orders.SB_GUID FROM orders "
					query = query & "Inner Join customers on orders.idCustomer = customers.idCustomer "
					query = query & "Inner Join SB_Orders on orders.idorder = SB_Orders.idorder "
					query = query & "WHERE orders.idorder = " & idorder
					set rsSB=Server.CreateObject("ADODB.Recordset")
					set rsSB=conntemp.execute(query)
					if err.number <> 0 then
						set rsSB=nothing
						call closeDb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in sb_CustViewSubs.asp: "&err.description) 
					end If
					if NOT rsSB.eof then
						   pcv_strCustEmail = rsSb("email")
						   pcv_strGUID = rsSb("SB_GUID")
					end if  
					set rsSB=nothing         
					%>
                    <tr>
                        <td>
                           <a href="OrdDetails.asp?id=<%response.write (scpre+int(IdOrder))%>"><%response.write (scpre+int(IdOrder))%></a>
                        </td>
                        <td>
                        	<%
							if pSubStartDate="" then
								response.Write("NA")
							else
								response.Write(showdateFrmt(pSubStartDate))
							end if
							%>
                        </td>
                        <td>
                           <%=pLinkID%>
                        </td>
                        <td>
                           <a href="https://www.subscriptionbridge.com/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=details" target="_blank"><%=pcv_strGUID%></a>
                        </td>
                        <td>
                            <%=scCurSign&money(pSubUnitPrice*pSubQty)%>
                        </td>
                        <td>
                            <%
							if pSubTrial = 1 then 
								if pSubTrialAmount > 0 Then
									Response.write scCurSign&money(pSubTrialAmount * pSubQty)
								Else						    
									Response.write "Free Trial"
								End if 
							Else
								Response.write "No Trial"						 
							End if 
                        	%>
                        </td>
                        <td nowrap><a href="https://www.subscriptionbridge.com/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=details" target="_blank">Manage</a></td>
					</tr>
					<%
					rstemp.movenext
			  	loop
				%>
			</table>
            
			<%
            if pResults<>"0" and iPageCount>1 Then
            %>
            <table class="pcCPcontent">
                <tr> 
                    <td> 
                        <%Response.Write("Page "& iPageCurrent & " of "& iPageCount & "<br />")%>
                        <%'Display Next / Prev buttons
                        if iPageCurrent > 1 then
                                'We are not at the beginning, show the prev button %>
                                 <a href="javascript:location='sb_ViewSubs.asp?idmain=<%=pcv_intIDMain%>&iPageCurrent=<%=iPageCurrent-1%>&curpage=<%=iPageCurrent%>';"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a>
                        <% end If
                        If iPageCount <> 1 then
                            For I=1 To iPageCount
                                If int(I)=int(iPageCurrent) Then %>
                                    <%=I%> 
                                <% Else %>
                                <a href="javascript:location='sb_ViewSubs.asp?idmain=<%=pcv_intIDMain%>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>&curpage=<%=iPageCurrent%>';" style="text-decoration:underline;"><%=I%></a>
                                <% End If %>
                            <% Next %>
                        <% end if %>
                        <% if CInt(iPageCurrent) <> CInt(iPageCount) then
                        'We are not at the end, show a next link %>
                        <a href="javascript:location='sb_ViewSubs.asp?idmain=<%=pcv_intIDMain%>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>&curpage=<%=iPageCurrent%>';"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a>
                        <% end If 
                        call closeDb()
                        %>
                    </td>
                </tr>
                <tr>
                    <td><hr></td>
                </tr>          
            </table>
            <% end if %>           
            
            
            
			<% 
            set rstemp = nothing
            call closeDb()
            %>
			</td>
		</tr>
		<tr>
			<td><hr></td>
		</tr>
		<tr> 
            <td>
            	<input type="button" value=" Main Menu " onClick="location.href='sb_Default.asp'">
                <input type="button" value=" View Packages " onClick="location.href='sb_ViewPackages.asp'">
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->