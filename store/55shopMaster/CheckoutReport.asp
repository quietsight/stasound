<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true %>
<% pageTitle="Drop-Off Reports" %>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--
	function chgWin(file,window) {
	msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=580,height=380');
	if (msgWindow.opener == null) msgWindow.opener = self;
}
//-->
</script>
<% 
'on error resume next
dim f, query, conntemp, rstemp, counter
dim strDateFormat
strDateFormat="mm/dd/yyyy"
if scDateFrmt="DD/MM/YY" then
	strDateFormat="dd/mm/yyy"
end if
counter=0

call openDb()

' count statistic registers
pcv_lastTotal=0
pcv_lastIncomp=0

if SQL_Format="1" then
	sDate=Day(Now)&"/"&Month(Now)&"/"&Year(now)-1
	eDate=Day(Now)&"/"&Month(Now)&"/"&Year(now)
else
	sDate=Month(Now)&"/"&Day(Now)&"/"&Year(now)-1
	eDate=Month(Now)&"/"&Day(Now)&"/"&Year(now)
end if

query=query&"Select count(*) as Total from orders "
if scDB="Access" then
	query=query&"WHERE orderdate>=#" & sDate & "# "
	query=query&"and orderdate<=#" & eDate & "# "
else
	query=query&"WHERE orderdate>='" & sDate & "' "
	query=query&"and orderdate<='" & eDate & "' "
end if
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	set rstemp = nothing
	call closedb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
end If
if not rstemp.eof then
	pcv_lastTotal=rstemp("Total")
end if 


query="Select count(*) as Incomplete from orders "
query=query&"WHERE orders.orderStatus=1 "
if scDB="Access" then
	query=query&"and orderdate>=#" & sDate & "# "
	query=query&"and orderdate<=#" & eDate & "# "
else
	query=query&"and orderdate>='" & sDate & "' "
	query=query&"and orderdate<='" & eDate & "' "
end if

'response.write query &"<br>"
'response.end
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	set rstemp = nothing
	call closedb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
end If
if not rstemp.eof then
	pcv_lastIncomp=rstemp("Incomplete")
end if 
set rstemp = nothing


viewyear=year(now())

if scDB="Access" then
	query="SELECT a.montha, a.Total, b.Incomplete, (b.Incomplete/a.Total)*100 as TotalPercent  "
else
	query="SELECT a.montha, a.Total, b.Incomplete, round(convert(float,b.Incomplete)/convert(float,a.Total)*100,0) as TotalPercent  "
end if
query=query&"from "
query=query&"( "
query=query&"Select count(orderdate) as Total, month(orderdate) AS montha from orders "
query=query&"WHERE year(orderdate)='" & viewyear & "' "
query=query&"GROUP BY month(orderdate) "
query=query&") a "
query=query&"left join "
query=query&"( "
query=query&"Select count(orderdate) as Incomplete , month(orderdate) AS monthb from orders "
query=query&"WHERE orders.orderStatus=1 "
query=query&"AND year(orderdate)='" & viewyear & "' "
query=query&"GROUP BY month(orderdate) "
query=query&") b "
query=query&"on a.montha = b.monthb;"
'query=query&"ORDER BY montha, monthb;"
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp = nothing
	call closedb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
end If
pcv_YearTotal=0
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2">
        <h2>Definitions</h2>
		<p>Read about incomplete orders, drop-offs, and conversions.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=469')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></p>
        <p>&nbsp;</p>
        <h2>Generate Reports</h2>
			<ul class="pcListIcon">
				<li><a href="#1">Drop-off by Date</a></li>
				<li><a href="#2">Drop-off by Product</a></li>
				<li><a href="#3">Drop-off by Customer Type</a></li>
				<li><a href="#4">Conversion Rate</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=441')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></li>
			</ul>
	</tr>
	<tr>
		<td colspan="2">
        <h2>Quick Summary: Monthly drop-off rates</h2>
		<%
		quantity=Cint(0)
		if rstemp.eof then
			quantity=Cint(0)
		%>
			<p>A drop-off report for the current year cannot be created as no incomplete orders exist. Please note that only incomplete orders that have been placed in the current year are included in drop-off reports.</p>
		<%
		else
			' creates array for chart
			cnt=month(now())
			ReDim arrValues(cnt)
			ReDim arrIncomp(cnt)			
			ReDim arrLabels(cnt)
			for lcnt=0 to cnt-1
				arrValues(lcnt)=0
				arrIncomp(lcnt)=0
				arrLabels(lcnt)= MonthName(lcnt+1,true)
			next
			do while not rstemp.eof 
				pTotalPercent=rstemp("TotalPercent")
				if not isNumeric(pTotalPercent) then
					pTotalPercent = 0
				end if
				pmonth=rstemp("montha")
				pIncomp=rstemp("Incomplete")
				if not isNumeric(pIncomp) then
					pIncomp = 0
				end if
				pTotal=rstemp("Total")
				pcv_YearIncomp=pcv_YearIncomp+Clng(pIncomp)
				pcv_YearTotal=pcv_YearTotal+Clng(pTotal)
				arrValues(pmonth-1)= Clng(pTotalPercent)
				arrIncomp(pmonth-1) = Clng(pIncomp)
				rstemp.movenext
			loop
			set rstemp=nothing
			Nspace=1
			%>
			<table width="100%" cellpadding="3" cellspacing="0">
			<tr>
				<td colspan="2" align="left" valign="middle" nowrap><b>Year To Date Monthly Breakdown</b></td>
			</tr>
			<%
			For k=lbound(arrValues) to Ubound(arrValues)-1
			%>
				<tr> 
					<td align="left" valign="middle" nowrap> 
					<%=arrLabels(k)%> total: <%=arrValues(k)%>%
					</td>
					<td width="100%" height="2" align="left" valign="middle">
					<%chartwidth=round(arrValues(k)/NSpace)
					if (chartwidth=0) or (chartwidth=1) then
						chartwidth=1
					end if%>
					<img src="images/pc_px.gif" height="10" width="<%=chartwidth%>" align="left" title="<%=arrIncomp(k)%>">
					</td>
				</tr>
			<% Next %>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td colspan="2" align="left" valign="middle" nowrap>Year To Date Total: <b><%=pcv_YearIncomp%> out of <%=pcv_YearTotal%> orders (<%=Clng((pcv_YearIncomp/pcv_YearTotal)*100)%>%)</b></td>
			</tr>
			<tr> 
				<td colspan="2" align="left" valign="middle" nowrap>Last 12 Months Total: <b><%=pcv_LastIncomp%> out of <%=pcv_LastTotal%> orders (<%=Clng((pcv_LastIncomp/pcv_LastTotal)*100)%>%)</b></td>
			</tr>
		</table>
		<% end if 'Have sales data%>
		
		</td>
	</tr>
</table>
<br />
<table class="pcCPcontent">
	<tr>
		<td colspan="2"><h2>Reports<a name="1"></a></h2></td>
	</tr>
	<tr> 
		<td width="60%" valign="top"> 
			<form action="dropoffReport.asp" name="date_form" target="_blank" class="pcForms">
			<% todayDate=Date() %>
			<p><b>View Drop-Off by Date</b></p>
			<p style="padding-top:10px;">
			<% Dim varMonth, varDay, varYear
			varMonth=Month(Date)
			varDay=Day(Date)
			varYear=Year(Date) 
			dim dtInputStrStart, dtInputStr
			dtInputStrStart=(varMonth&"/01/"&varYear)
			if scDateFrmt="DD/MM/YY" then
				dtInputStrStart=("01/"&varMonth&"/"&varYear)
			end if
			dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
			if scDateFrmt="DD/MM/YY" then
				dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
			end if
			%>
			From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
			To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
			</p>
			<p style="padding-top:10px;">
			Country:
			<%
			query="SELECT CountryCode,countryName FROM countries ORDER BY countryName ASC"
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=conntemp.execute(query)
			if err.number <> 0 then
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
			end If
			%>
			<select name="CountryCode">
			<option value="" selected>-- All countries --</option>
			<%
			do while not rstemp.eof
				pCountryCode2=rstemp("CountryCode")%>
				<option value="<%response.write pCountryCode2%>" <%if pCountryCode2=scShipFromPostalCountry then%>selected<%end if %>><%response.write rstemp("countryName")%></option>
			<%
			rstemp.movenext
			loop
			set rstemp = nothing
			%>
			</select>
			</p>
			<p style="padding-top:10px;">
			<input type="submit" value="Search" name="submit" class="submit2">
			</p>
			</form>
		</td>
		<td width="40%" valign="top">Specify a date range to view all drop-offs in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></td>
	</tr>
<tr>
<td colspan="2"><hr><a name="2"></a></td>
</tr>
	<tr> 
		<td width="60%" valign="top"> 
			<form action="PrddropoffReport.asp" name="prdsales_form" target="_blank" class="pcForms">
			<% todayDate=Date() %>
			<p><b>View Drop-Off by Product</b></p>
			<p align="left" style="padding-top:10px;">
			From:	<input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
			To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
			</p>
			<p style="padding-top:10px;">
			Select a product (only products included on incomplete orders are shown):
			</p>
			<%
			query="SELECT distinct ProductsOrdered.IDproduct,products.description,products.sku FROM ProductsOrdered,products,Orders WHERE  products.removed=0 AND products.active=-1 AND orders.idorder=ProductsOrdered.idorder and ProductsOrdered.IDproduct=products.IDproduct and orders.orderStatus=1 ORDER BY products.description ASC"
			
			set rstemp=Server.CreateObject("ADODB.Recordset")
			set rstemp=conntemp.execute(query)
			if err.number <> 0 then
				set rstemp=nothing
				call closeDb()
				response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
			end If
			if NOT rstemp.eof then
				prdArray = rstemp.getRows()
				intCount=ubound(prdArray,2)
			else
				intCount=0
			end if
			set rstemp = nothing			
			%>
			<p style="padding-top:10px;">
			<select name="IDProduct">
			<option value="" selected>-- All products --</option>
			<% 
			if intCount>0 then
				for i=0 to intCount
				%>
					<option value="<%=prdArray(0,i)%>"><%=prdArray(1,i)%> (<%=prdArray(2,i)%>)</option>
				<% 
				next 
			end if
			%>
			</select>
			</p>
			<p style="padding-top:10px;">
			<input type="submit" value="Search" name="submit" class="submit2">
			</p>
			</form>
			</td>
		<td width="40%" valign="top">Specify a product and date range to view all drop-offs in that period. Only products for which an incomplete order has been recorded are shown in the drop-down. <b>Note</b>: You must enter both dates in the format <font color="#000099"><%=strDateFormat%></font></td>
	</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
	<tr> 
		<td width="60%"> 
			<form action="custDropoffReport.asp" name="aff_form" target="_blank" class="pcForms">
			<a name="3"></a>
			<p><b>View Drop-Off by Customer Type</b></p>
			<p align="left" style="padding-top:10px;">
			From:	<input class="textbox" type="text" size="10" name="FromDate" value="<%=dtInputStrStart%>">
			To:	<input class="textbox" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">
			</p>
			<p style="padding-top:10px;">
			<input type="submit" value="Search" name="submit" class="submit2">
			</p>
			</form>
		</td>
		<td width="40%" valign="top">Specify a date range to view all drop-offs by customer type in that period. <b>Note</b>: You must enter both dates in the format <font color="#000099"><%=strDateFormat%>.</font></td>
	</tr>

<% if scVersion>="3" then%>
	<tr>
		<td colspan="2"><hr><a name="4"></a></td>
	</tr>
		<tr> 
			<td width="60%"> 
				<a name="aff"></a>
				<form action="custConversionReport.asp" name="aff_form" target="_blank" class="pcForms">
				<p><b>View Customer Conversion Rate</b>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=441')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></p>
				<p align="left" style="padding-top:10px;">
				From:	<input class="textbox" type="text" size="10" name="FromDate" value="<%=dtInputStrStart%>">
				To:	<input class="textbox" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">
				</p>
				<p style="padding-top:10px;">
				<input type="submit" value="Search" name="submit" class="submit2">
				</p>
				</form>
			</td>
			<td width="40%" valign="top">Specify a date range to view all customer conversion rates in that period. <b>Note</b>: You must enter both dates in the format <font color="#000099"><%=strDateFormat%>.</font> Only customers which are created AFTER upgrading to v3.0 will be shown as "New Customers" within the date range.</td>
		</tr>
<% end if %>
</TABLE>
<% call closeDb() %>
<!--#include file="AdminFooter.asp"-->