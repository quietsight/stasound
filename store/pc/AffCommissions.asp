<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="AffLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<%
'*******************************
' Check if store is ON or OFF
'*******************************
If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If
%>
<!--#include file="header.asp"-->
	<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td> 
				<h1><%=dictLanguage.Item(Session("language")&"_AffCom_1")%></h1>
			</td>
		</tr>
		<tr>
			<td>
				<table class="pcShowContent">
					<%
					' Load affiliate ID
					affVar=session("pc_idaffiliate")
					if not validNum(affVar) then
						response.redirect "AffiliateLogin.asp"
					end if
					
					Dim rs,query 
					Dim tempId
					tempId=0
					
					' Our Connection Object
					Dim con
					Set con=CreateObject("ADODB.Connection")
					con.Open scDSN 
		
					' Choose the records to display	
					query="SELECT * FROM Orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12))"
					query=query&" AND idaffiliate="& affVar &" ORDER BY orders.orderDate desc"
					' Our Recordset Object
	
					Set rs=CreateObject("ADODB.Recordset")
					rs.CursorLocation=adUseClient
					rs.Open query, scDSN , 3, 3
				
					' If the returning recordset is not empty
					If rs.EOF Then %>
						<tr> 
							<td colspan="7">
								<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_AffCom_2")%></div>
							</td>
						</tr>
					<% Else
						dim conntemp
						call opendb()
						
						query="SELECT SUM(affiliatePay) AS AfftotalSum FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND idAffiliate=" & affVar
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=connTemp.execute(query)
						AffTotalSum=rstemp("AfftotalSum")
						if AffTotalSum<>"" then
						else
							AffTotalSum=0
						end if
					
						query="SELECT SUM(pcAffpay_Amount) AS AfftotalPaid FROM pcAffiliatesPayments WHERE pcAffpay_idAffiliate=" & affVar
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=connTemp.execute(query)
						AffTotalPaid=rstemp("AfftotalPaid")
						if AffTotalPaid<>"" then
						else
							AffTotalPaid=0
						end if
						
						CurrentBalance=AffTotalSum-AffTotalPaid
						%>
						<tr> 
							<td colspan="7"><%=dictLanguage.Item(Session("language")&"_AffCom_3")%><%=scCursign%><%=money(CurrentBalance)%></td>
						</tr>
						<tr> 
							<td colspan="7"><%=dictLanguage.Item(Session("language")&"_AffCom_4")%></td>
						</tr>
						<tr> 
							<td colspan="7">
								<table id="AutoNumber1" class="pcShowContent">
									<tr>
										<th width="20%"><%=dictLanguage.Item(Session("language")&"_AffCom_5")%></th>
										<th width="20%"><div align="right"><%=dictLanguage.Item(Session("language")&"_AffCom_6")%></div></th>
										<th width="60%"><%=dictLanguage.Item(Session("language")&"_AffCom_7")%></th>
									</tr>
									<%query="SELECT pcAffpay_idpayment, pcAffpay_Amount, pcAffpay_PayDate, pcAffpay_Status FROM pcAffiliatesPayments WHERE pcAffpay_idAffiliate=" & affVar & " ORDER BY pcAffpay_PayDate DESC;"
									set rstemp=server.CreateObject("ADODB.RecordSet")
									set rstemp=connTemp.execute(query)
									if rstemp.eof then%>
										<tr>
											<td colspan="3">
												<div class="pcErrorMessage">
												<%=dictLanguage.Item(Session("language")&"_AffCom_8")%>
												</div>
											</td>
										</tr>
									<%else
										do while not rstemp.eof
											IDpayment=rstemp("pcAffpay_idpayment")
											PaidAmount=rstemp("pcAffpay_Amount")
											PaidDate=rstemp("pcAffpay_PayDate")
											'// Format Date
											PaidDate=ShowDateFrmt(PaidDate)
											PaidStatus=rstemp("pcAffpay_Status")%>
											<tr>
												<td><%=PaidDate%></td>
												<td><div align="right"><%=scCurSign%>&nbsp;<%=money(PaidAmount)%></div></td>
												<td><%=PaidStatus%></td>
											</tr>
											<%rstemp.MoveNext
										loop%>
										<tr>
											<td><div align="right"><%=dictLanguage.Item(Session("language")&"_AffCom_9")%></div></td>
											<td><div align="right"><%=scCurSign%>&nbsp;<%=money(AffTotalPaid)%></div></td>
											<td>&nbsp;</td>
										</tr>
									<%end if%>
								</table>
							</td>
						</tr>
						<tr>
							<td colspan="7" class="pcSpacer"></td>
						</tr>
						<tr> 
							<td colspan="7"><%=dictLanguage.Item(Session("language")&"_AffCom_10")%><%=rs.RecordCount%></td>
						</tr>
						<tr>
							<td colspan="7" class="pcSpacer"></td>
						</tr>
						<tr> 
							<th nowrap><%=dictLanguage.Item(Session("language")&"_AffCom_11")%></th>
							<th nowrap><%=dictLanguage.Item(Session("language")&"_AffCom_12")%></th>
							<th nowrap><%=dictLanguage.Item(Session("language")&"_AffCom_15")%></th>
							<th nowrap><%=dictLanguage.Item(Session("language")&"_AffCom_16")%></th>
							<th nowrap><%=dictLanguage.Item(Session("language")&"_AffCom_17")%></th>
							<th nowrap><%=dictLanguage.Item(Session("language")&"_AffCom_18")%></th>
							<th nowrap><%=dictLanguage.Item(Session("language")&"_AffCom_13")%></th>
						</tr>
						<% 
						gTotalsales=0
						gTotaltaxes=0
						gTotalOrder=0
						aTotalShip=0
						gTotalTax=0
						gTotalcomm=0
						do until rs.EOF
							gSubOrder=0
							gSubTax=0
							gSubShip=0
							gSubCom=0

							intIdOrder=rs("idOrder")
							intIdCustomer=rs("idcustomer")
							dtOrderDate=rs("orderDate")
							dblAffiliatePay=rs("affiliatePay")
							porderdetails=rs("details")
				
							'Calculate "NET" Order Amount
							ptotal=rs("total")
							gSubOrder=rs("total")
							ptaxAmount=rs("taxAmount")
							ptaxDetails=rs("taxDetails")
							pord_VAT=rs("ord_VAT")
							gSubTax=ptaxAmount+pord_VAT
							pshipmentDetails=rs("shipmentDetails")
							Postage=0
							serviceHandlingFee=0
							shipping=split(pshipmentDetails,",")
							if ubound(shipping)>1 then
								if NOT isNumeric(trim(shipping(2))) then
								else
									Postage=trim(shipping(2))
									if ubound(shipping)=>3 then
										serviceHandlingFee=trim(shipping(3))
										if NOT isNumeric(serviceHandlingFee) then
											serviceHandlingFee=0
										end if
									else
										serviceHandlingFee=0
									end if
								end if
							end if
							gSubShip=Postage
							ppaymentDetails=trim(rs("paymentDetails"))
							payment = split(ppaymentDetails,"||")
							PayCharge=0
							If ubound(payment)>=1 then					
								If payment(1)="" then
									PayCharge=0
								else
									PayCharge=payment(1)
								end If
							End if

							PrdSales=ptotal
							PrdSales=PrdSales-postage
							PrdSales=PrdSales-serviceHandlingFee
							PrdSales=PrdSales-PayCharge
							
							gSubOrder=gSubOrder-postage


							pdiscountDetails=rs("discountDetails")
							pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
							if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
								pcv_CatDiscounts="0"
							end if
							
							if (instr(pdiscountDetails,"- ||")>0) or (pcv_CatDiscounts>"0")  then
								if instr(pdiscountDetails,",") then
									DiscountDetailsArry=split(pdiscountDetails,",")
									intArryCnt=ubound(DiscountDetailsArry)
								else
									intArryCnt=0
								end if
									
								dim discounts, discountType 
													
								discount=0
								for k=0 to intArryCnt
									if instr(pTempDiscountDetails,"- ||") then
										discounts = split(pTempDiscountDetails,"- ||")
										tdiscount = discounts(1)
									else
										tdiscount=0
									end if
									discount=discount+tdiscount
								Next
								PrdSales=PrdSales+discount+pcv_CatDiscounts
							end if
							
							if pord_VAT>0 then
								PrdSales=PrdSales-pord_VAT
							else
								if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
									PrdSales=PrdSales-ptaxAmount
								else 
									taxArray=split(ptaxDetails,",")
									for i=0 to (ubound(taxArray)-1)
										taxDesc=split(taxArray(i),"|")
										PrdSales=PrdSales-taxDesc(1)
										gSubTax=gSubTax+taxDesc(1)
									next 
								end if
							end if

							gSubOrder=gSubOrder-gSubTax
							
							gTotalsales=gTotalsales + PrdSales
							gTotalOrder=gTotalOrder+gSubOrder
							gTotalShip=gTotalShip+gSubShip
							gTotalTax=gTotalTax+gSubTax
							%>
							<tr> 
								<% '// Format Date
								dtOrderDate=rs("orderDate")
								dtOrderDate=ShowDateFrmt(dtOrderDate) %>
								<td nowrap><%=dtOrderDate%></td>
								<td nowrap>#<%=cdbl(scpre)+cdbl(rs("idOrder"))%></td>
								<td align="right" nowrap><%=scCurSign&money(PrdSales)%></td>
								<td align="right" nowrap><%=scCurSign&money(gSubOrder)%></td>
								<td align="right" nowrap><%=scCurSign&money(gSubShip)%></td>
								<td align="right" nowrap><%=scCurSign&money(gSubTax)%></td>
								<td align="right" nowrap><%=scCurSign&money(rs("affiliatePay"))%></td>
							</tr>
							<% gTotalcomm=gTotalcomm + rs("affiliatePay") %>
							<% rs.MoveNext
						loop
					End If %>
					<tr> 
					  <td colspan="7"><hr></td>
					</tr>
					<tr> 
					  <td colspan="7"><div align="right"><b><%=dictLanguage.Item(Session("language")&"_AffCom_14")%></b></div></td>
					</tr>
					<tr> 
					  <td colspan="7"><div align="right"><b><%=scCurSign&money(gTotalcomm)%></b></div></td>
					</tr>
					<tr> 
					  <td colspan="7"><hr></td>
					</tr>
					<tr> 
					  <td colspan="7"><a href="javascript:history.go(-1)"><img src="<%=rslayout("back")%>"></a></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>
<%	' Done. Now release Objects
con.Close
Set con=Nothing
Set rs=Nothing
%>
<%call closedb()%>
<!--#include file="footer.asp"-->