<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Adjust Google Analytics Statistics" %>
<% Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<% Session.LCID = 1033 %>
<%
on error resume next

'// GOOGLE ANALYTICS
'// E-commerce transaction tracking: order refunds and cancellations
'// http://www.google.com/support/analytics/bin/answer.py?answer=27203&topic=7282

'// If Google Analytics is not active, redirect.
if trim(scGoogleAnalytics)="" or IsNull(scGoogleAnalytics) then
	response.redirect "msg.asp?message=44"
end if

dim query, conntemp, rstemp, pOID, pcGAtransaction, paffiliateName, paffiliateCompany, ptaxAmount, ptaxDetails, pshipmentDetails, ptotal, pGetItems, pcItemArray, pcItemArrayCount, iCount, itemInfo, pExistsInLog


'// STEP 2 - Get order data and prompt customer to edit
'// Get order ID



		pOID=trim(request.QueryString("idOrder"))
			if pOID="" then
				pOID=trim(request.Form("OrderNumber"))
				if pOID<>"" then
					pOID=(int(pOID)-scpre)
				end if
			end if
			
		'// Query string passed when a transaction already exists in the log but the store manager
		'// wants to post it again.
		pNextAction=trim(request.QueryString("confirm"))
			
		'// See if there is a message to display
		msg=trim(request.QueryString("msg"))
			
		'// IF the order id is there, show details for it.
		'// Otherwise, show the list of orders to choose from.
		IF pOID<>"" and pNextAction<>"YES" THEN ' 0
		
		
				'// CHECK LOGS - START
				'// Check to see if this order has already been adjusted
					pExistsInLog=0
					
					Dim sDSNFile
					sDSNFile = "gaLog.dsn"
					
					' Let's now dynamically retrieve the current directory
					Dim sScriptDir
					sScriptDir = Request.ServerVariables("SCRIPT_NAME")
					sScriptDir = StrReverse(sScriptDir)
					sScriptDir = Mid(sScriptDir, InStr(1, sScriptDir, "/"))
					sScriptDir = StrReverse(sScriptDir)
					
					' Time to construct our dynamic DSN
					Dim sPath, sDSN
					sPath = Server.MapPath(sScriptDir) & "\GAlogs\"
					sDSN = "FileDSN=" & sPath & sDSNFile & _
								 ";DefaultDir=" & sPath & _
								 ";DBQ=" & sPath & ";"
					
					Dim newConn
					Set newConn = Server.CreateObject("ADODB.Connection")
					newConn.Open sDSN

					query = "SELECT ORDERNUMBER,DATE,TRANSACTIONINFO,ITEMINFO FROM gaLog.txt WHERE ORDERNUMBER="&(int(pOID)+scpre)
					set rs = newConn.execute(query)
					
					'Print out the contents of our recordset
					If not rs.EOF then
						pExistsInLog=1
						Response.Write "<div style='padding:15px'>This order has already been adjusted. The following information was logged when the trasaction was posted (<a href='GAlogs/galog.txt' target='_blank'>view log file</a>)."
						Do While Not rs.EOF
							pcArrayOrderInfo=split(rs("TRANSACTIONINFO"),"|")
							Response.Write "<br><br>"
							Response.Write "<div><strong>Order Number</strong>: " & rs("ORDERNUMBER") & "</div>"
							Response.Write "<div><strong>Adjustment posted to Google Analytics on</strong>: " & rs("DATE") & "</div>"
							Response.Write "<div><strong>Adjustment Amount</strong>: " & pcArrayOrderInfo(3) & "</div>"
							replaceString="UTM:I|"&(int(pOID)+scpre)&"|"
							itemInfo=replace(rs("ITEMINFO"),replaceString,"")
							itemInfo=replace(itemInfo,"|","&nbsp;&nbsp;&nbsp;")
							Response.Write "<strong>Adjustment Items</strong> (Part Number, Name, Category, Unit Price, Units):<br>" & itemInfo & "<br>"
							rs.MoveNext
							response.write "<br><br>If you would like to post a new adjustment for this order, <a href='pcGA_refund.asp?idorder="&pOID&"&confirm=YES"&"'>click here</a>.<br><br>To look for another order, <a href='pcGA_refund.asp'>click here</a></div>"
						Loop
					End If
					
					'Close our recordset and connection
					rs.close
					set rs = nothing
					newConn.close
					set newConn = nothing
				
				'// CHECK LOGS - END
		END IF ' 0
		
		IF (pOID<>"" and pNextAction="YES") OR (pOID<>"" and pExistsInLog<>1) THEN ' 1
				
				pGetItems = 1
				
				call openDb()
				
				'// Google Analytics
				'// (a) GENERATE ORDER INFO LINE
				'// (b) GENERATE ITEM INFO LINES
				
				'// (a) - START
				
					'// Get order info from db
				
					query="SELECT total, city, state, stateCode, CountryCode, shipmentDetails, idAffiliate, taxAmount, taxDetails, ord_VAT FROM orders WHERE idOrder=" & pOID
					set rs=Server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
					
					if err.number <> 0 then
						set rs=nothing
						pGetItems = 0
						call closeDb()
					end If
					
					if rs.eof then
						set rs=nothing
						pGetItems = 0
						call closeDb()
					end if
								
					IF pGetItems<>0 THEN ' 2 - If database errors, skip everything
				
						'// Transaction line variables per Google Analytics documentation
						'// [order-id]  Your internal unique order id number  
						'// [affiliation]  Optional partner or store affilation  
						'// [total]  Total dollar amount of the transaction  
						'// [tax]  Tax amount of the transaction  
						'// [shipping]  The shipping amount of the transaction  
						'// [city]  City to correlate the transaction with  
						'// [state/region]  State or province  
						'// [country]  Country
						
						pIdOrder = pOID
						pTotal = rs("total")
						
						'// Gather affiliate information
						pidAffiliate=rs("idaffiliate")
							If pidaffiliate>"1" then
								query="SELECT affiliateName, affiliateCompany FROM affiliates WHERE idAffiliate =" & pidAffiliate
								Set rsTemp=Server.CreateObject("ADODB.Recordset")
								Set rsTemp=connTemp.execute(query)
								paffiliateName = rsTemp("affiliateName")
								paffiliateCompany = rsTemp("affiliateCompany")
									if trim(paffiliateCompany)<>"" then
										paffiliateName = paffiliateName & "(" & paffiliateCompany & ")"
									end if
								paffiliateName = replace(paffiliateName,"|","-")
								Set rsTemp = nothing
							else
								paffiliateName = "N/A"
							end if
														
						'// Gather tax information to determine total tax amount
						ptaxAmount=rs("taxAmount")
						ptaxDetails=rs("taxDetails")
						pord_VAT=rs("ord_VAT")
						if pord_VAT>0 then
							ptaxAmount=pord_VAT
						else
							if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
								ptaxAmount=ptaxAmount
							else
								taxArray=split(ptaxDetails,",")
								ptaxTotal=0
								for i=0 to (ubound(taxArray)-1)
									taxDesc=split(taxArray(i),"|")
									ptaxTotal=ptaxTotal+taxDesc(1)
								next 
								ptaxAmount=ptaxTotal
							end if
						end if
						
						'// Gather shipping information to determine total shipping amount
						pshipmentDetails=rs("shipmentDetails")
							pTotalShipping=0
							shipping=split(pshipmentDetails,",")
							if ubound(shipping)>1 then
								if NOT isNumeric(trim(shipping(2))) then
									pTotalShipping=0
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
									if serviceHandlingFee<>0 then
										pTotalShipping=CDbl(Postage)+CDbl(serviceHandlingFee)
									else
										pTotalShipping=CDbl(Postage)
									end if
								end if
							end if
								
						'// Gather order location information
						pcity=rs("city")
							pcity = replace(pcity,"|","-")
						pstate=rs("state")
						pstateCode=rs("stateCode")
							if trim(pstateCode)="" then
								pstateCode=pstate
							end if
						pCountryCode=rs("CountryCode")
						
						set rs = nothing
					
					END IF ' 2 - End If database errors, skip everything
						
				'// (a) - END
				
				'// (b) - START
												
					'// Gather item info from database
					
					query="SELECT ProductsOrdered.idProduct, ProductsOrdered.quantity, ProductsOrdered.unitPrice, products.description, products.sku FROM ProductsOrdered, products WHERE ProductsOrdered.idProduct=products.idProduct AND ProductsOrdered.idOrder=" & pOID
					set rs=Server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
					
					if err.number <> 0 then
						set rs=nothing
						pGetItems = 0
						call closeDb()
					end If
					
					if rs.eof then
						set rs=nothing
						pGetItems = 0
						call closeDb()
					end if
							
					IF pGetItems<>0 THEN ' 3 - If database errors, skip everything
							
						'// Item line variables  
						'// [order-id]  Your internal unique order id number (should be same as transaction line)  
						'// [sku/code]  Product SKU code  
						'// [product name]  Product name or description  
						'// [category]  Category of the product or variation  
						'// [price]  Unit-price of the product  
						'// [quantity]  Quantity ordered
						
						pcItemArray=rs.getRows()
						pcItemArrayCount=ubound(pcItemArray,2)
						set rs = nothing
						
					END IF ' 3 - End If database errors, skip everything
				
				'// STEP 2 - END
				
				call closeDb()
				
				'// Write the hidden form
				%>
				<form action="pcGA_refundSubmit.asp" class="pcForms" method="post" name="GArefund">
					<input type="hidden" name="idOrder" value="<%=pIdOrder%>">
					<input type="hidden" name="itemCount" value="<%=pcItemArrayCount%>">
					<table class="pcCPcontent">
						<tr>
							<td colspan="4" class="pcCPspacer">
							<% if msg<>"" then %>
								<div class="pcCPmessage"><%=msg%></div>
							<% end if %>
							</td>
						</tr>
						<tr>
							<th colspan="4">Order #: <%=(scpre+int(pIdOrder))%> - General Information&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=315')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
						</tr>
						<tr>
							<td colspan="4" class="pcCPspacer"></td>
						</tr>
						<tr style="background-color:#F1F1F1;">
							<td></td>
							<td>Original</td>
							<td align="center">Adjustment</td>
							<td></td>
						</tr>
						<tr>
							<td nowrap="nowrap">Order Total:</td>
							<td><%=ptotal%></td>
							<td nowrap="nowrap"><%=scCurSign%>&nbsp;<input type="text" value="<%=(ptotal*-1)%>" name="total" size="8"></td>
							<td rowspan="3" valign="top" style="padding-left: 15px;"><div>These are the amounts that will be posted to Google Analytics to offset the original transaction. For a <strong>full refund/cancellation</strong>, leave 'as is' (posted amount fully offsets the original amount).</div><div style="padding-top: 6px">Note that the original transaction <u>is not removed</u> from the system. Also note that if the <u>date of the refund/cancellation</u> is different from the date of the original transaction, some reports might include one and not the other.</div></td>
						</tr>
						<tr>
							<td>Taxes:</td>
							<td><%=ptaxAmount%></td>
							<td nowrap="nowrap"><%=scCurSign%>&nbsp;<input type="text" value="<%=(ptaxAmount*-1)%>" name="taxes" size="8"></td>
						</tr>
						<tr>
							<td>Shipping:</td>
							<td><%=pTotalShipping%></td>
							<td nowrap="nowrap"><%=scCurSign%>&nbsp;<input type="text" value="<%=(pTotalShipping*-1)%>" name="shipping" size="8"></td>
						</tr>
					</table>
					<table class="pcCPcontent">
						<tr>
							<td colspan="4" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="4">Order #: <%=(scpre+int(pIdOrder))%> - Item Information&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=315')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th></th>
						</tr>
						<tr>
							<td colspan="4" class="pcCPspacer"></td>
						</tr>
						<tr style="background-color:#F1F1F1;">
							<td nowrap="nowrap">Item Name</td>
							<td nowrap="nowrap">Units Purchased</td>
							<td nowrap="nowrap">Units <strong>Returned/Refunded</strong></td>
							<td nowrap="nowrap">Unit Price</td>
						</tr>
						
						<%
						iCount = 0
						For iCount=0 to pcItemArrayCount
							
							pIdProduct = pcItemArray(0,iCount)
							pName = pcItemArray(3,iCount)
								pName = replace(pName,"|","-")
							pUnitPrice = pcItemArray(2,iCount)
							pQuantity = pcItemArray(1,iCount)
						%>
						
						<tr>
							<td><strong><%=pName%></strong></td>
							<td><%=pQuantity%></td>
							<td><input type="text" value="<%=(pQuantity*-1)%>" name="quantity<%=pIdProduct%>" size="10"></td>
							<td><%=pUnitPrice%></td>
						</tr>
						<%
						next
						%>
						<tr>
							<td colspan="4"><hr></td>
						</tr>
						<tr>
							<td colspan="4">
								<input type="submit" class="submit2" value="Post Adjustments">&nbsp;
								<input type="button" value="Back" onClick="Javascript:history.back()">
							</td>
						</tr>
					</table>
				</form>
		
		<%
		END IF ' 1
		
		IF pOID="" THEN ' 2
		%>
		
			<%
			'// *****************************************
			'// ORDERS - Show a list of orders
			'// *****************************************
			
			Const iPageSize=10
			
			Dim iPageCurrent
			
			if request.querystring("iPageCurrent")="" then
				iPageCurrent=1
			else
				iPageCurrent=Request.QueryString("iPageCurrent")
			end if
			
			'sorting order
			Dim strORD
			
			strORD=request("order")
			if strORD="" then
				strORD="orderDate DESC, idOrder"
			End If
			
			strSort=request("sort")
			if strSort="" Then
				strSort="DESC"
			End If
			
			query1=""
			
			OType=request("OType")
			if OType="" then
				OType="0"
			end if
			if (OType<>"0") then
				query1= " orderstatus=" & OType
				else
				query1= " orderstatus>2"
			end if
			
			pcv_PayType=request("PayType")
			if (pcv_PayType<>"") then
				query1= query1 & " AND pcOrd_PaymentStatus=" & pcv_PayType
			end if
			
			err.number=0
			FromDate=request("fromdate")
			PassFromDate=FromDate
			ToDate=request("todate")
			PassToDate=ToDate
			
			if FromDate<>"" then
				if scDateFrmt="DD/MM/YY" then
					DateVarArray=split(FromDate,"/")
					FromDate=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
				else
					if SQL_Format="1" then
						FromDate=day(FromDate) & "/" & month(FromDate) & "/" & year(FromDate)
					else
						FromDate=month(FromDate) & "/" & day(FromDate) & "/" & year(FromDate)
					end if
				end if
			else
				call opendb()
				query="SELECT TOP 1 orders.orderDate FROM orders WHERE orders.orderStatus>1 ORDER BY orderDate ASC;"
				set rstemp=Server.CreateObject("ADODB.Recordset") 
				set rstemp=conntemp.execute(query)
				if NOT rstemp.eof then
					FromDate=rstemp("orderDate")
					if scDateFrmt="DD/MM/YY" then
						FromDate=day(FromDate)&"/"&month(FromDate)&"/"&Year(FromDate)
						PassFromDate=FromDate
					else
						if SQL_Format="1" then
							FromDate=day(FromDate) & "/" & month(FromDate) & "/" & year(FromDate)
						else
							FromDate=month(FromDate) & "/" & day(FromDate) & "/" & year(FromDate)
						end if
					end if
				end if
				call closedb()
			end if
			
			if ToDate<>"" then
				if scDateFrmt="DD/MM/YY" then
					DateVarArray2=split(ToDate,"/")
					ToDate=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
				else
					if SQL_Format="1" then
						ToDate=day(ToDate) & "/" & month(ToDate) & "/" & year(ToDate)
					else
						ToDate=month(ToDate) & "/" & day(ToDate) & "/" & year(ToDate)
					end if
				end if
			else
				if SQL_Format="1" then
					ToDate=day(date()) & "/" & month(date()) & "/" & year(date())
				else
					ToDate=month(date()) & "/" & day(date()) & "/" & year(date())
				end if
			end if
			
			if trim(request("fromdate"))="" then
				dtFromDate=Date()-13
				dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
				if SQL_Format="1" then
					FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
				else
					FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
				end if
				dtToDate=Date()
				dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
				if SQL_Format="1" then
					ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
				else
					ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
				end if
				if scDateFrmt="DD/MM/YY" then
					dtFromDateStr=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
					dtToDateStr=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
				end if
				PassFromDate=dtFromDateStr
				PassToDate=dtToDateStr
			end if
			
			if (FromDate<>"") and (IsDate(FromDate)) then
				if scDB="Access" then
					query1= query1 & " AND orderDate>=#" & FromDate & "#"
				else
					query1= query1 & " AND orderDate>='" & FromDate & "'"
				end if
			end if
			
			if (ToDate<>"") and (IsDate(ToDate)) then
				if scDB="Access" then
					query1= query1 & " AND orderDate<=#" & ToDate & "#"
				else
					query1= query1 & " AND orderDate<='" & ToDate & "'"
				end if
			end if
			
			call openDb()
			
			' Choose the records to display
			Dim srcVar
				SqlVar="SELECT orders.idOrder, orders.idCustomer, orders.paymentCode, orders.orderstatus, orders.orderDate, orders.total, orders.pcOrd_PaymentStatus, customers.name, customers.lastName, customers.customerCompany FROM orders, customers WHERE orders.idCustomer=customers.idCustomer AND " & query1 & " ORDER BY "& strORD &" "& strSort
			%>
			
			<% 
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			
			rstemp.CursorLocation=adUseClient
			rstemp.CacheSize=iPageSize
			rstemp.PageSize=iPageSize
			rstemp.Open SqlVar, conntemp
			
			if err.number <> 0 then
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
			end If
			
			rstemp.MoveFirst
			' get the max number of pages
			Dim iPageCount
			iPageCount=rstemp.PageCount
			If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
			If iPageCurrent < 1 Then iPageCurrent=1
			
			' set the absolute page
			rstemp.AbsolutePage=iPageCurrent
			if rstemp.eof then 
				presults="0"
			else
				%>
				<table class="pcCPcontent">
					<tr> 
						<td width="40%" valign="top">
							<b>
							<%
							if PassFromDate<>"" then %>
								From: <%=PassFromDate%>
							<%end if%>
							&nbsp;
							<%if PassToDate="" then
								PassToDate=date()
								if scDateFrmt="DD/MM/YY" then
									PassToDate=day(PassToDate)&"/"&month(PassToDate)&"/"&Year(PassToDate)
								end if %>
							<% end if%>
							To: <%=PassToDate%></b>
							<br>
							<% ' Showing total number of pages found and the current page number
							Response.Write "Displaying Page <b>" & iPageCurrent & "</b> of <b>" & iPageCount & "</b><br>"
							Response.Write "Total Records Found : <b>" & rstemp.RecordCount & "</b>" %>
						</td>
						<td width="60%" valign="top">
						Locate the order from the list below and click on &quot;Select&quot; or enter the order number here and press the &quot;Start&quot; button.
						<form action="pcGA_refund.asp" method="post" class="pcForms">
							Order Number: <input type="text" value="" name="OrderNumber" size="8">
							<input type="submit" value="Start" class="submit2">
							<input type="button" value="View Logs" onClick="document.location.href='pcGA_logs.asp'">
						</form>			
						</td>
					</tr>
				</table>
			<% end if %>

			<table class="pcCPcontent">
				<tr>
					<td colspan="7" align="center" nowrap>&nbsp;</td>
				</tr>
				<tr> 
					<th align="center" nowrap><a href="pcGA_refund.asp?FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=I%>&order=orderstatus&sort=ASC"><img src="images/sortasc_blue.gif" alt="Sort Ascending"></a><a href="pcGA_refund.asp?FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=I%>&order=orderstatus&sort=DESC"><img src="images/sortdesc_blue.gif" alt="Sort Descending"></a></th>
					<th align="center" nowrap><a href="pcGA_refund.asp?FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=I%>&order=pcOrd_PaymentStatus&sort=ASC"><img src="images/sortasc_blue.gif" alt="Sort Ascending"></a><a href="pcGA_refund.asp?FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=I%>&order=pcOrd_PaymentStatus&sort=DESC"><img src="images/sortdesc_blue.gif" alt="Sort Descending"></a></th>
					<th align="center" nowrap><a href="pcGA_refund.asp?FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=I%>&order=orderDate&sort=ASC"><img src="images/sortasc_blue.gif" alt="Sort Ascending"></a><a href="pcGA_refund.asp?FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate%>&iPageCurrent=<%=I%>&order=orderDate&sort=DESC"><img src="images/sortdesc_blue.gif" alt="Sort Descending"></a>&nbsp;Date</th>
					<th nowrap>ID</th>
					<th nowrap>Customer</th>
					<th nowrap>Total</th>
					<th width="3%" nowrap></th>
				</tr>
				<tr>
					<td colspan="7" class="pcCPspacer"></td>
				</tr>
				<% 
				Dim mcount
				mcount=0
				If rstemp.EOF Then %>
				<tr>
					<td colspan="7" align="center">
						<div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20"> No Results Found</div>
					</td>
				</tr>
				<% Else
				' Showing relevant records
				Dim strCol
					strCol="#E1E1E1" 
				Dim rcount, i, x
				
				For i=1 To rstemp.PageSize
					pidOrder=rstemp("idOrder")
					pidCustomer=rstemp("idCustomer")
					ppaymentCode=rstemp("paymentCode")
					porderstatus=rstemp("orderstatus")
					porderDate=rstemp("orderDate")
					ptotal=rstemp("total")
					pcv_PaymentStatus=rstemp("pcOrd_PaymentStatus")
					if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
						pcv_PaymentStatus=0
					end if
					pfName=rstemp("name")
					plName=rstemp("lastName")
					pCustomerCompany=rstemp("customerCompany")
					if trim(pCustomerCompany)<>"" then
						pCustomerName=pfName & " " & plName & " (" & pCustomerCompany & ")"
						else
						pCustomerName=pfName & " " & plName
					end if			
					rcount=i
					If currentPage > 1 Then
						For x=1 To (currentPage - 1)
							rcount=10 + rcount
						Next
					End If
					
					'// DeActivate Sections for Google Checkout
					pcv_strDeactivateStatus=0						
					if ppaymentCode="Google" then
						pcv_strDeactivateStatus=1
					end if	
					
					If Not rstemp.EOF Then 
						If strCol <> "#FFFFFF" Then
							strCol="#FFFFFF"
						Else 
							strCol="#E1E1E1"
						End If
						mcount=mcount+1 %>
						<tr bgcolor="<%= strCol %>"> 
							<td align="center" valign="top"> 
							<% select case porderstatus
								case "1"
									response.Write "<img src=""images/purpledot.gif"" alt=""Incomplete"">"
								case "2"
									response.write "<img src=""images/bluedot.gif"" alt=""Pending"">" 
								case "3"
									response.write "<img src=""images/yellowdot.gif"" alt=""Processed"">" 
								case "4"
									response.write "<img src=""images/greendot.gif"" alt=""Shipped"">" 
								case "5"
									response.write "<img src=""images/reddot.gif"" alt=""Canceled"">" 
								case "6"
									response.write "<img src=""images/orangedot.gif"" alt=""Return"">" 
								case "7"
									response.write "<img src=""images/7dot.gif"" alt=""Partially Shipped"">"
								case "8"
									response.write "<img src=""images/8dot.gif"" alt=""Shipping"">"
								case "9"
									response.write "<img src=""images/9dot.gif"" alt=""Partially Return"">"
								case "10"
									response.write "<img src=""images/greendot.gif"" alt=""Delivered"">" 
								case "11"
									response.write "<img src=""images/reddot.gif"" alt=""Will Not Deliver"">" 
								case "12"
									response.write "<img src=""images/greendot.gif"" alt=""Archived"">" 
								case "13"
									response.write "<img src=""images/darkgreendot.gif"" alt=""Refund"">"
							end select %>
							</td>
							<td align="center" valign="top"> 
							<% select case pcv_PaymentStatus
							case "0"
								response.write "<img src=""images/blueflag.gif"" width=""15"" height=""13"" alt=""Pending"">" 
							case "1"
								response.write "<img src=""images/yellowflag.gif"" width=""15"" height=""13"" alt=""Authorized"">" 
							case "2"
								response.write "<img src=""images/greenflag.gif"" width=""15"" height=""13"" alt=""Paid"">" 
							case "6"
								response.write "<img src=""images/darkgreenflag.gif"" width=""15"" height=""13"" alt=""Refunded"">" 
							case "8"
								response.write "<img src=""images/redflag.gif"" width=""15"" height=""13"" alt=""Voided"">" 
							end select %>
							</td>
							<td align="left" valign="top" bgcolor="<%= strCol %>"><%=ShowDateFrmt(porderDate)%></td>
							<td><a href="ordDetails.asp?id=<%=pidOrder%>"><%=(scpre+int(pidOrder))%></a></td>
							<td><%=pCustomerName%></td>
							<td nowrap="nowrap"><%=scCurSign&money(ptotal)%></td>
							<td nowrap="nowrap"><a href="pcGA_refund.asp?idOrder=<%=pidorder%>">Select &gt;&gt;</a></td>
						</tr>
					<% rstemp.MoveNext
					End If
				Next%>
			<%End If %>
			</table>        

			<% if pResults<>"0" Then %>
			<table class="pcCPcontent">
				<tr>
					<td><hr></td>
				</tr>
				<tr> 
					<td> 
						<%Response.Write("Page "& iPageCurrent & " of "& iPageCount & "<br />")%>
						<%'Display Next / Prev buttons
						if iPageCurrent > 1 then
								'We are not at the beginning, show the prev button %>
								 <a href="pcGA_refund.asp?FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a>
						<% end If
						If iPageCount <> 1 then
							For I=1 To iPageCount
								If I=iPageCurrent Then %>
						<%=I%> 
						<% Else %>
						<a href="pcGA_refund.asp?FromDate=<%=PassFromDate %>&ToDate=<%=PassToDate %>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>"><%=I%></a>
						<% End If %>
						<% Next %>
						<% end if %>
						<% if CInt(iPageCurrent) <> CInt(iPageCount) then
						'We are not at the end, show a next link %>
						<a href="pcGA_refund.asp?FromDate=<%=PassFromDate %>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a>
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

			<table class="pcCPcontent" style="width:auto;">
				<tr> 
					<td><b>Advanced Filters</b></td>
				</tr>
				<tr> 
					<td>Filter orders by date and status.</td>
				</tr>
				<tr> 
					<td>
					<form action="pcGA_refund.asp" name="advsearch" class="pcForms">
						<table class="pcCPcontent">
							<tr>
								<td align="right">Date From:</td>
								<td nowrap="nowrap"><input type="text" name="fromdate" value="<%=PassFromDate%>" size="10"> To: <input type="text" name="todate" value="<%=PassToDate%>" size="10">	<i>(<%=lcase(scDateFrmt)%>)</i>
								</td>
							</tr>
							<tr>
								<td align="right" valign="top" nowrap="nowrap">Order Status:</td>
								<td>
									<select name="otype">
										<option value="0" <%if OType="0" then%>selected<%end if%>>All</option>
										<option value="2" <%if OType="2" then%>selected<%end if%>>Pending</option>
										<option value="3" <%if OType="3" then%>selected<%end if%>>Processed</option>
										<option value="7" <%if OType="7" then%>selected<%end if%>>Partially Shipped</option>
										<option value="8" <%if OType="8" then%>selected<%end if%>>Shipping</option>
										<option value="4" <%if OType="4" then%>selected<%end if%>>Shipped</option>
										<option value="5" <%if OType="5" then%>selected<%end if%>>Canceled</option>
										<option value="9" <%if OType="9" then%>selected<%end if%>>Partially Return</option>
										<option value="6" <%if OType="6" then%>selected<%end if%>>Return</option>	
										<option value="13" <%if OType="13" then%>selected<%end if%>>Refund</option>	
										<% if GOOGLEACTIVE<>0 then %>
										<option value="10" <%if OType="10" then%>selected<%end if%>>Delivered</option>
										<option value="11" <%if OType="11" then%>selected<%end if%>>Will Not Deliver</option>
										<option value="12" <%if OType="12" then%>selected<%end if%>>Archived</option>
										<% end if %>						
									</select>
								</td>
							</tr>
							<tr>
								<td align="right" valign="top" nowrap="nowrap">Payment Status:</td>
								<td>
									<select name="PayType">
										<option value="" <%if pcv_PayType="" then%>selected<%end if%>>All</option>
										<option value="0" <%if pcv_PayType="0" then%>selected<%end if%>>Pending</option>
										<option value="1" <%if pcv_PayType="1" then%>selected<%end if%>>Authorized</option>
										<option value="2" <%if pcv_PayType="2" then%>selected<%end if%>>Paid</option>
										<% if GOOGLEACTIVE<>0 then %>
										<option value="3" <%if pcv_PayType="3" then%>selected<%end if%>>Declined</option>
										<option value="4" <%if pcv_PayType="4" then%>selected<%end if%>>Cancelled</option>
										<option value="5" <%if pcv_PayType="5" then%>selected<%end if%>>Cancelled By Google</option>							
										<option value="7" <%if pcv_PayType="7" then%>selected<%end if%>>Charging</option>
										<% end if %>
										<option value="6" <%if pcv_PayType="6" then%>selected<%end if%>>Refunded</option>
										<option value="8" <%if pcv_PayType="8" then%>selected<%end if%>>Voided</option>
									</select>
								</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2" align="center">
									<input type="submit" name="B1" value="Search Orders" class="submit2">
									&nbsp;
									<input type="button" name="Button" value="Back" onClick="location='invoicing.asp'">
								</td>
							</tr>
						</table>
					</form>
					</td>
				</tr>
			</table>

	<%
	END IF ' 2
	%>
<!--#include file="AdminFooter.asp"-->