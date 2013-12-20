<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Locate an order" %>
<% Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/javascripts/pcDateFunctions.js"-->
<%Dim connTemp,rs,query

call opendb()
%>
<!--#include file="pcCharts.asp"-->
<%dim strDateFormat
strDateFormat="mm/dd/yyyy"
if scDateFrmt="DD/MM/YY" then
	strDateFormat="dd/mm/yyyy"
end if
%>
<script language="JavaScript">
<!--
	function Validate_Dates(theForm)
	{
	
		if (theForm.FromDate.value == "")
		{
			alert("Please enter From Date and try again.");
			theForm.FromDate.focus();
			return (false);
		}
		
		if (theForm.ToDate.value == "")
		{
			alert("Please enter To Date and try again.");
			theForm.ToDate.focus();
			return (false);
		}
		
		if (isDate(theForm.FromDate.value,"<%=strDateFormat%>","From Date")==false)
		{
			theForm.FromDate.focus()
			return false
		}
		
		if (isDate(theForm.ToDate.value,"<%=strDateFormat%>","To Date")==false)
		{
			theForm.ToDate.focus()
			return false
		}
		
		if (CompareDates(theForm.FromDate,theForm.ToDate,"From <= To")==false)
		{
			alert("From Date should be less than To Date.")
			theForm.ToDate.focus()
			return false
		}
	return (true);
	}
//-->
</script>
<script>
	/*
	Clear default form value script- By Ada Shimar (ada@chalktv.com)
	Featured on JavaScript Kit (http://javascriptkit.com)
	Visit javascriptkit.com for 400+ free scripts!
	*/
	function clearText(thefield){
	if (thefield.defaultValue==thefield.value)
	thefield.value = ""
	} 
	
function CalPop(sInputName)
{
	window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
}
</script>
	<table class="pcCPcontent" style="background-image:url(images/pcv4_icon_search100.gif); background-repeat: no-repeat;">
		<tr>
			<td>
				<table class="pcCPcontent" style="width: auto; margin-left: 100px;">
                	<tr>
                    	<td>Select one of the following:</td>
                    </tr>
					<tr>
						<td valign="top">
							<ul>
								<li><a href="#today">View today's orders</a></li>
								<li><a href="resultsAdvancedAll.asp?B1=View+All&dd=1">View recent orders</a></li>
								<li><a href="resultsAdvancedAll.asp?B1=View+All&dd=1&OType=1">View incomplete orders</a></li>
							</ul>
						</td>
						<td valign="top">
							<ul>
								<li><a href="#datestatus">Filter orders by date and status</a></li>
								<li><a href="#keyword">Filter orders by keyword</a></li>
								<li><a href="viewCusta.asp">Filter orders by customer</a></li>
							</ul>
						</td>
						<td valign="top">
							<ul>
								<li><a href="#payment">Filter orders by payment type</a></li>
								<li><a href="#registry">Filter orders by Gift Registry</a></li>
								<li><a href="resultsAdvancedAll.asp?B1=View+All&pcIntArchived=1">View archived orders</a></li>
							</ul>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr> 
			<th>Quick Summary in last 30 days</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
		<td>
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%" valign="top">
					<div id="chartOrdStatus30days" style="height:330px; "></div>
				</td>
				<td width="50%" valign="top">
					<div id="chartOrder30days" style="height:250px; "></div>
				</td>
			</tr>
			</table>
			<%
				call pcs_Gen30daysALLOrdersCharts("chartOrder30days",0)
				call pcs_OrdStatus30Days("chartOrdStatus30days")
			%>
		</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th><a name="today">&nbsp;</a>Orders received today: <%= ShowMonthFrmt((now))%></th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td>  
				<%			 
				
				Todaydate=Date()
				if SQL_Format="1" then
					Todaydate=Day(Todaydate)&"/"&Month(Todaydate)&"/"&Year(Todaydate)
				else
					Todaydate=Month(Todaydate)&"/"&Day(Todaydate)&"/"&Year(Todaydate)
				end if
										
				if scDB="Access" then
					query="SELECT idorder, orderDate, total FROM orders WHERE orderStatus>1 AND orderStatus<5 AND orderDate=#"&Todaydate&"#" 
				else
					query="SELECT idorder, orderDate, total FROM orders WHERE ((orderStatus>1 AND orderStatus<5) OR (orderStatus>6 AND orderStatus<9) OR (orderStatus=10 OR orderStatus=12)) AND orderDate='"&Todaydate&"'" 
				end if 
						
				Set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				if rs.eof then%>
				No orders have been placed today yet.
				<% else %>
					<ul>
						<% While Not rs.eof
							pTotal=rs("total")
							pIdOrder=rs("idOrder")
							pIdOrder=scpre + int(pIdOrder) %>
							<li><%response.write "Order #: "&pIdOrder&" - Total: "&scCurSign & money(pTotal)%> - <a class="resultslink" href="Orddetails.asp?id=<%=rs("idorder")%>" onFocus="if(this.blur)this.blur()">View details &gt;&gt;</a></li>
							<%rs.MoveNext%>
						<%Wend%>
					</ul>
				<% end if 
				Set rs=Nothing %>
			</td>
		</tr>
		<tr>
			<td><a name="datestatus">&nbsp;</a></td>
		</tr>
		<tr> 
			<th>Seach Orders by Date &amp; Order and/or Payment Status</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td>Enter a date range and select the order and/or payment status from the drop-menu below:</td>
		</tr>
		<tr> 
			<td>
				<form action="resultsAdvancedAll.asp?" name="advsearch" align="center" class="pcForms" onSubmit="return Validate_Dates(this)">
					<table class="pcCPcontent" style="width:auto;">
						<tr>
							<td>Date:</td>
							<% dim dtDateFrom, dtDateTo
							if scDateFrmt="DD/MM/YY" then
								if day(date())-13 > 1 then
									dtDateFrom=day(date())-13 & "/" & month(date()) & "/" & year(date())
								else
									dtDateFrom="01/" & month(date()) & "/" & year(date())
								end if
								dtDateTo=day(date()) & "/" & month(date()) & "/" & year(date())
							else
								if day(date())-13 > 1 then
									dtDateFrom=month(date()) & "/" & day(date())-13 & "/" & year(date())
								else
									dtDateFrom=month(date()) & "/01/" & year(date())
								end if
								dtDateTo=month(date()) & "/" & day(date()) & "/" & year(date())
							end if %>
							<td  valign="top">From: 
							<input type="text" name="FromDate" value="<%=dtDateFrom%>" size="10">&nbsp;<a href="javascript:CalPop('document.advsearch.FromDate');"><img SRC="../Calendar/icon_Cal.gif" border="0" align="absbottom"></a>
							 To: 
							<input type="text" name="ToDate" value="<%=dtDateTo%>" size="10">&nbsp;<a href="javascript:CalPop('document.advsearch.ToDate');"><img SRC="../Calendar/icon_Cal.gif" border="0" align="absbottom"></a>
							</td>
						</tr>
						<tr>
							<td>Status:</td>
							<td>
							<select name="otype">
								<option value="0" selected>All</option>
								<option value="2">Pending</option>
								<option value="3">Processed</option>
								<option value="7">Partially Shipped</option>
								<option value="8">Shipping</option>
								<option value="4">Shipped</option>
								<option value="5">Canceled</option>
								<option value="9">Partially Returned</option>
								<option value="6">Returned</option>
								<option value="1">Incomplete</option>
								<% if GOOGLEACTIVE=-1 then %>
								<option value="10" >Declined</option>
								<option value="12" >Archived</option>
								<% end if %>
							</select>
							</td>
						</tr>
						<tr>
							<td>Payment Status:</td>
							<td>
							<select name="PayType">
								<option value="" selected>All</option>
								<option value="0">Pending</option>
								<option value="1">Authorized</option>
								<option value="2">Paid</option>
								<option value="6">Refunded</option>
								<option value="8">Voided</option>
								<% if GOOGLEACTIVE=-1 then %>
								<option value="3">Declined</option>
								<option value="4">Cancelled</option>
								<option value="5">Cancelled By Google</option>
								<option value="7">Charging</option>
								<% end if %>
							</select>
							</td>
						</tr>
                        <tr>
                            <td align="right" valign="top" nowrap>Archived Orders Only:</td>
                            <td>
                                <input type="checkbox" name="pcIntArchived" value="1" class="clearBorder">
                            </td>
                        </tr>
						<tr>
							<td colspan="2"><input type="submit" name="B1" value="Search Orders" class="submit2"></td>
						</tr>
					</table>
				</form>
			</td>
		</tr>
		<tr>
			<td><a name="keyword">&nbsp;</a></td>
		</tr>
		<tr> 
			<th>Search Orders by Keyword</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td>
				Choose a filter from the drop down menu, then specify a keyword that the filter should use. For example, to list all orders that included &quot;widgets&quot;, choose &quot;Product&quot; as a filter and enter &quot;widget&quot; as a keyword. Click <b>Search</b> to begin the search.
			</td>
		</tr>
		<tr> 
			<td valign="top" align="left">
			<script language="JavaScript">
			<!--
			
			function Form1_Validator(theForm)
			{
			if (theForm.advquery.value=="Please enter a keyword")
			{
			theForm.advquery.value="";	
			}
			return (true);
			}
			//-->
			</script> 
			<form action="resultsAdvanced.asp?" name="advsearch2" onSubmit="return Form1_Validator(this)" class="pcForms">
				<select class="select2" name="TypeSearch" size="1">
					<option value="idOrder">Order ID</option>
                    <option value="orderCode">Order Code</option>
					<% if GOOGLEACTIVE=-1 then %>
					<option value="GoogleOrderID">Google Order ID</option>
					<% end if %>
					<option value="details">Product</option>
					<option value="shipmentDetails">Shipping Type</option>
					<option value="stateCode">State Code</option>
					<option value="CountryCode">Country Code</option>
				</select>
                &nbsp;
				<input class="textbox"  type="text" size="25" name="advquery" value="Please enter a keyword" onFocus="clearText(this)">
                <div style="margin-top: 10px; margin-bottom: 10px;">
				<input type="submit" name="B1" value="Search" class="submit2">
                </div>
			</form></td>
		</tr>
		<tr>
			<td><a name="payment">&nbsp;</a></td>
		</tr>
		<tr> 
			<th>Search Orders by Payment Type</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 				
			<td>Choose a payment type from the drop-down meny and click <b>Search</b> to list order submitted using that payment type.</td>
		</tr>
		<tr> 
			<td> 
			<form action="resultsAdvanced.asp" name="PayType" class="pcForms" method="get">
            <input type="hidden" name="TypeSearch" value="payment">
			<% query="SELECT DISTINCT (paymentDesc), idPayment FROM payTypes ORDER BY paymentDesc ASC"
			Set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query) %>
			<select class="select" name="advquery" size="1">
				<%
				Do While Not rs.EOF
					strPaymentDesc=rs("paymentDesc")
					intIdPayment=rs("idPayment") %>
					<option value="<%=strPaymentDesc%>"><%=strPaymentDesc %></option>
					<% rs.movenext					
				loop %>
				<%
				Set rs=Nothing
				%>
			</select>
			<div style="margin-top: 10px; margin-bottom: 10px;">
			<input type="submit" name="Submit" value="Search" class="submit2">
            </div>
			</form>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"><a name="registry">&nbsp;</a></td>
		</tr>        
	<%
	'// Orders by Gift Registry
	
	query="SELECT pcEv_IDEvent,pcEv_Name,pcEv_Type,pcEv_IDCustomer FROM pcEvents ORDER BY pcEv_Name ASC;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)	
	if not rs.eof then
	%>
		<tr> 
			<th>Search Orders by Gift Registry</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 				
			<td>View all the orders that have been placed from a specific Gift Registry.</td>
		</tr>
		<tr> 
			<td> 
			<form action="resultsAdvanced.asp?" name="ordersByRegistry" onSubmit="return Form1_Validator(this)" class="pcForms">
            	<input type="hidden" name="TypeSearch" value="registry">
            	<select name="pcIntRegistryID">
                	<option selected value="0">Select an existing Gift Registry</option>
                    <%
					do while not rs.eof
					gIDEvent=rs("pcEv_IDEvent")
					gName=rs("pcEv_Name")
					gType=rs("pcEv_Type")
					pcv_IdCustomer=rs("pcEv_IDCustomer")
					
					'Check to see if the customer exists in the DB
					query="SELECT name FROM customers WHERE idcustomer="&pcv_IdCustomer&";"
					set rsTemp=server.CreateObject("ADODB.RecordSet")
					set rsTemp=conntemp.execute(query)
					if rsTemp.EOF then
						pcv_CustomerCheck="NO"
					end if
					if pcv_CustomerCheck<>"NO" then
						query="SELECT customers.name, customers.lastName FROM customers WHERE idcustomer="&pcv_IdCustomer&";"
						set rsTemp=conntemp.execute(query)
						pcv_strCustName=rsTemp("name") & " " & rsTemp("lastName")
					else
						pcv_strCustName="<font style='color:#FF0000;'>Customer has<br>been deleted</font>"
					end if
					%>
                    <option value="<%=gIDEvent%>"><%=gName%> - <%=gType%> (<%=pcv_strCustName%>)</option>
                    <%
					rs.movenext
					loop
					%>
                </select>
			<div style="margin-top: 10px; margin-bottom: 10px;">
			<input type="submit" name="Submit" value="Search" class="submit2">
            </div>
			</form>
          </td>
       </tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	<%	
    end if
	set rstemp=nothing
	set rs=nothing
	call closeDb()
    %>

</table>
<!--#include file="AdminFooter.asp"-->