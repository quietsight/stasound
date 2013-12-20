<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Copy custom field to other products" %>
<% nav=request("nav")
if nav="bto" then
	Section="services"
else
	Section="products"
end if %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, conntemp, rs
idproduct=request.querystring("idproduct")

call openDb()
query="SELECT description, xfield1, xfield2, xfield3 FROM products WHERE idproduct="&idproduct
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=conntemp.execute(query)
productName=rs("description")
xfield1=rs("xfield1")
xfield2=rs("xfield2")
xfield3=rs("xfield3")
mytest=false
%>
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<table class="pcCPcontent">
	<tr>
		<td>
			<h2>Product: <strong><%=productName%></strong></h2>
			Select the custom field that you would like to copy to other products:
            
				<form name="choices" action="applyCustomFields2.asp?action=next" method="post" class="pcForms">
				<input type="hidden" name="idproduct" value="<%=idproduct%>">
				<%query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & idproduct & ";"
				set rs=connTemp.execute(query)
				if not rs.eof then%>
				<table>
					<tr> 
						<td colspan="3"><b>Custom Search Fields</b></td>
					</tr>
					<tr> 
					<td nowrap="nowrap" colspan="2">Text to display</td>
					<td nowrap="nowrap">Value</td>
					</tr>
					<%pcArr=rs.getRows()
					set rs=nothing
					intCount=ubound(pcArr,2)
					mytest=true
					For i=0 to intCount	%>									
					<tr> 
						<td>
						<input type="radio" name="CustField" value="custom<%=pcArr(2,i)%>" <%if session("CustFieldCopy")="custom" & pcArr(2,i) then%>checked<%end if%> class="clearBorder"></td>
						<td width="275" nowrap><%=pcArr(1,i)%></td>
						<td width="100%"><%=pcArr(3,i)%></td>
					</tr>
					<%Next%>
				</table>
				<hr>
				<% end if
				set rs=nothing%>
				<table>
					<tr>
						<td colspan="2"><b>Custom Input Fields</b></td>
					</tr>
					<tr> 
						<td colspan="2">Text to display</td>
					</tr>
								
				<%
						query="SELECT xfield FROM xfields WHERE idxfield="&xfield1
						idxfield=xfield1
						set rs=Server.CreateObject("ADODB.Recordset")     
						set rs=conntemp.execute(query)
						if rs.eof then
							xfield1="Field not assigned"
						else
							xfield1=rs("xfield")
						end if
						set rs=nothing
														
						if xfield1<>"Field not assigned" then
							mytest=true
					%>
								
						<tr> 
							<td><input type="radio" name="CustField" value="xfield1" <%if session("CustFieldCopy")="xfield1" then%>checked<%end if%> class="clearBorder"></td>
							<td width="100%"><%=xfield1%>&nbsp;</td>
						</tr>
				<% end if %>
								
				<%
						query="SELECT xfield FROM xfields WHERE idxfield="&xfield2
						idxfield=xfield2
						set rs=Server.CreateObject("ADODB.Recordset")     
						set rs=conntemp.execute(query)
						if rs.eof then
							xfield2="Field not assigned"
						else
							xfield2=rs("xfield")
						end if
						set rs=nothing
							
						if xfield2<>"Field not assigned" then
							mytest=true
				%>
					<tr> 
						<td><input type="radio" name="CustField" value="xfield2" <%if session("CustFieldCopy")="xfield2" then%>checked<%end if%> class="clearBorder"></td>
						<td width="100%"><%=xfield2%></td>
					</tr>
				<% end if %>
								
				<%
						query="SELECT xfield FROM xfields WHERE idxfield="&xfield3
						idxfield=xfield3
						set rs=Server.CreateObject("ADODB.Recordset")     
						set rs=conntemp.execute(query)
						if rs.eof then
							xfield3="Field not assigned"
						else
							xfield3=rs("xfield")
						end if
						set rs=nothing
						
						if xfield3<>"Field not assigned" then
							mytest=true
					%>
					<tr> 
						<td><input type="radio" name="CustField" value="xfield3" <%if session("CustFieldCopy")="xfield3" then%>checked<%end if%> class="clearBorder"></td>
						<td width="100%"><%=xfield3%></td>
					</tr>
				<% end if %>
				</table>
				<%
					set rs=nothing
					call closeDb()
				%>
				<br /><br />
				<%if mytest=false then%>
					<div class="pcCPmessage">No custom fields have been assigned to this product. Please select another product.</div>
				<%end if%>
					
					
				<p align="center">
				<%if mytest=true then%>&nbsp;
					<input type="submit" name="submit" value=" Next Step >> " class="submit2">&nbsp;
				<%end if%>
                	<input type="button" name="back" value="Back" onClick="location='AdminCustom.asp?idproduct=<%=idproduct%>';">
				</p>
				</form>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->