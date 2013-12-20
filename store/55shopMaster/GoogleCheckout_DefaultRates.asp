<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Google Checkout - Advanced Shipping Setup: Default Rates" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="AdminHeader.asp"-->

<%
Dim connTemp, query, rs

if request.form("Submit")<>"" then		
		
		call opendb()
		query="SELECT idshipservice, servicePriority, serviceDescription, serviceDefaultRate FROM shipService WHERE serviceActive=-1 ORDER BY servicePriority;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		do until rs.eof 
		
			a=rs("idshipservice")
			b=rs("idshipservice")
			d=Request(a)
			if d="" then
				d=0
			end if			
			query="UPDATE shipService SET serviceDefaultRate="& d &" WHERE idshipservice=" & b
			set rs2=server.CreateObject("ADODB.RecordSet")
			set rs2=connTemp.execute(query)
			set rs2 = nothing
			
		rs.moveNext
		loop 
		set rs = nothing
		call closedb()	
		

	response.redirect "GoogleCheckout_DefaultRates.asp"
else
%>
<form name="data" action="GoogleCheckout_DefaultRates.asp" method="post" class="pcForms">
	<table class="pcCPcontent">
		<tr> 
			<td colspan="2"><strong><u>Note</u></strong>: these "default" rates are <u>only shown</u> to the customer if Google cannot communicate with your store after the checkout process begins. "Default" shipping rates are not a ProductCart requirement, but rather they are required by Google Checkout. Here you can set your own default rates, which will override the rates determined by ProductCart. Make sure to click on the &quot;Save&quot; button when you are done.</td>
		</tr>
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>	
        <tr> 
            <th>Service</th> 
            <th>Default Rate</th>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
		<% 
		set rs=server.CreateObject("ADODB.RecordSet")
		call opendb()
		query="SELECT idshipservice, servicePriority, serviceDescription, serviceDefaultRate FROM shipService WHERE serviceActive=-1 ORDER BY servicePriority, serviceDescription;"
		set rs=connTemp.execute(query)
		do until rs.eof 
			a=rs("idshipservice")
			b=rs("serviceDescription")
			c=rs("serviceDefaultRate")
			if isNULL(c)=True OR c="" then
				c=0
			end if
			%>
			<tr> 
				<td width="30%" nowrap><%=b%>:</td> 
				<td width="70%" align="left">
				<%=scCurSign%> <input name="<%=a%>" type="text" value="<%=money(c)%>">
				</td>
			</tr>
			<%
		rs.moveNext
		loop 
		call closedb()
		%>			
		<tr>
			<td class="pcCPspacer" colspan="2"><hr></td>
		</tr>
		<tr>           
			<td colspan="2">
				<input  name="submit" type="submit" onClick="return setHidden(this.form)" value="Save" class="submit2"/>
				&nbsp;
				<input type="button" value="Shipping Options Summary" onClick="location.href='viewShippingOptions.asp'">				
			</td>
		</tr>
	</table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->